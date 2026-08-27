//! Strict, borrowed table-model discovery facts.
//!
//! The Keynote table-listing path needs only the identity/display name and
//! dimensions of a `TST.TableModelArchive`.  This codec deliberately keeps
//! that boundary smaller than the complete generated model: the strict raw
//! pass validates every field and preserves the caller-owned bytes, while the
//! two existing private Buffa lazy views provide a generated-free parity
//! check for the selected scalar fields.  No generated value or collection
//! crosses this module boundary.
//!
//! Unknown fields are accepted only when their framing is canonical.  Unknown
//! groups are consumed and retained by the caller-owned source bytes, subject
//! to the same finite nesting/field/work limits as known fields.  This is the
//! policy used by the existing table-name and table-header projections.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The strict wire pass intentionally precedes the Buffa parity check."
)]

use std::{fmt, str};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_numbers_names_generated::LitchiIwaProjection as names_projection;
use crate::buffa_numbers_table_header_settings_generated::LitchiIwaProjection as dimensions_projection;

const TABLE_ID_FIELD: u32 = 1;
const TABLE_STYLE_FIELD: u32 = 3;
const BASE_DATA_STORE_FIELD: u32 = 4;
const TABLE_ROWS_FIELD: u32 = 6;
const TABLE_COLUMNS_FIELD: u32 = 7;
const TABLE_NAME_FIELD: u32 = 8;
const DEFAULT_ROW_HEIGHT_FIELD: u32 = 16;
const DEFAULT_COLUMN_WIDTH_FIELD: u32 = 17;
const BODY_CELL_STYLE_FIELD: u32 = 18;
const HEADER_ROW_STYLE_FIELD: u32 = 19;
const HEADER_COLUMN_STYLE_FIELD: u32 = 20;
const FOOTER_ROW_STYLE_FIELD: u32 = 21;
const BODY_TEXT_STYLE_FIELD: u32 = 24;
const HEADER_ROW_TEXT_STYLE_FIELD: u32 = 25;
const HEADER_COLUMN_TEXT_STYLE_FIELD: u32 = 26;
const FOOTER_ROW_TEXT_STYLE_FIELD: u32 = 27;
const MAX_RECURSION: u32 = 64;
const MAX_FIELD_NUMBER: u32 = 0x1fff_ffff;
const BUFFA_PARITY_PASSES: usize = 2;

const fn doubled_at_least_one(value: usize) -> usize {
    match value.checked_mul(2) {
        Some(0) => 1,
        Some(value) => value,
        None => usize::MAX,
    }
}

const REQUIRED_FIELDS_MASK: u32 = (1 << TABLE_ID_FIELD)
    | (1 << TABLE_STYLE_FIELD)
    | (1 << BASE_DATA_STORE_FIELD)
    | (1 << TABLE_ROWS_FIELD)
    | (1 << TABLE_COLUMNS_FIELD)
    | (1 << TABLE_NAME_FIELD)
    | (1 << DEFAULT_ROW_HEIGHT_FIELD)
    | (1 << DEFAULT_COLUMN_WIDTH_FIELD)
    | (1 << BODY_CELL_STYLE_FIELD)
    | (1 << HEADER_ROW_STYLE_FIELD)
    | (1 << HEADER_COLUMN_STYLE_FIELD)
    | (1 << FOOTER_ROW_STYLE_FIELD)
    | (1 << BODY_TEXT_STYLE_FIELD)
    | (1 << HEADER_ROW_TEXT_STYLE_FIELD)
    | (1 << HEADER_COLUMN_TEXT_STYLE_FIELD)
    | (1 << FOOTER_ROW_TEXT_STYLE_FIELD);

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum KnownFieldKind {
    String,
    Varint,
    Bool,
    Fixed64,
    Bytes,
    RepeatedVarint,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct KnownField {
    kind: KnownFieldKind,
    wire: u32,
    required: bool,
    repeated: bool,
    name: &'static str,
}

const fn known_field(number: u32) -> Option<KnownField> {
    let field = match number {
        TABLE_ID_FIELD => KnownField {
            kind: KnownFieldKind::String,
            wire: 2,
            required: true,
            repeated: false,
            name: "TST.TableModelArchive.table_id",
        },
        TABLE_STYLE_FIELD => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: true,
            repeated: false,
            name: "TST.TableModelArchive.table_style",
        },
        BASE_DATA_STORE_FIELD => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: true,
            repeated: false,
            name: "TST.TableModelArchive.base_data_store",
        },
        TABLE_ROWS_FIELD => KnownField {
            kind: KnownFieldKind::Varint,
            wire: 0,
            required: true,
            repeated: false,
            name: "TST.TableModelArchive.number_of_rows",
        },
        TABLE_COLUMNS_FIELD => KnownField {
            kind: KnownFieldKind::Varint,
            wire: 0,
            required: true,
            repeated: false,
            name: "TST.TableModelArchive.number_of_columns",
        },
        TABLE_NAME_FIELD => KnownField {
            kind: KnownFieldKind::String,
            wire: 2,
            required: true,
            repeated: false,
            name: "TST.TableModelArchive.table_name",
        },
        DEFAULT_ROW_HEIGHT_FIELD => KnownField {
            kind: KnownFieldKind::Fixed64,
            wire: 1,
            required: true,
            repeated: false,
            name: "TST.TableModelArchive.default_row_height",
        },
        DEFAULT_COLUMN_WIDTH_FIELD => KnownField {
            kind: KnownFieldKind::Fixed64,
            wire: 1,
            required: true,
            repeated: false,
            name: "TST.TableModelArchive.default_column_width",
        },
        BODY_CELL_STYLE_FIELD => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: true,
            repeated: false,
            name: "TST.TableModelArchive.body_cell_style",
        },
        HEADER_ROW_STYLE_FIELD => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: true,
            repeated: false,
            name: "TST.TableModelArchive.header_row_style",
        },
        HEADER_COLUMN_STYLE_FIELD => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: true,
            repeated: false,
            name: "TST.TableModelArchive.header_column_style",
        },
        FOOTER_ROW_STYLE_FIELD => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: true,
            repeated: false,
            name: "TST.TableModelArchive.footer_row_style",
        },
        BODY_TEXT_STYLE_FIELD => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: true,
            repeated: false,
            name: "TST.TableModelArchive.body_text_style",
        },
        HEADER_ROW_TEXT_STYLE_FIELD => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: true,
            repeated: false,
            name: "TST.TableModelArchive.header_row_text_style",
        },
        HEADER_COLUMN_TEXT_STYLE_FIELD => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: true,
            repeated: false,
            name: "TST.TableModelArchive.header_column_text_style",
        },
        FOOTER_ROW_TEXT_STYLE_FIELD => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: true,
            repeated: false,
            name: "TST.TableModelArchive.footer_row_text_style",
        },
        5 => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.provider",
        },
        9 => KnownField {
            kind: KnownFieldKind::Varint,
            wire: 0,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.number_of_header_rows",
        },
        10 => KnownField {
            kind: KnownFieldKind::Varint,
            wire: 0,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.number_of_header_columns",
        },
        11 => KnownField {
            kind: KnownFieldKind::Varint,
            wire: 0,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.number_of_footer_rows",
        },
        12 => KnownField {
            kind: KnownFieldKind::Bool,
            wire: 0,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.header_rows_frozen",
        },
        13 => KnownField {
            kind: KnownFieldKind::Bool,
            wire: 0,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.header_columns_frozen",
        },
        14 => KnownField {
            kind: KnownFieldKind::Varint,
            wire: 0,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.number_of_hidden_rows",
        },
        15 => KnownField {
            kind: KnownFieldKind::Varint,
            wire: 0,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.number_of_hidden_columns",
        },
        22 => KnownField {
            kind: KnownFieldKind::Bool,
            wire: 0,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.table_name_enabled",
        },
        23 => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.origin_offset",
        },
        28 => KnownField {
            kind: KnownFieldKind::Varint,
            wire: 0,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.preset_index",
        },
        29 => KnownField {
            kind: KnownFieldKind::Bool,
            wire: 0,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.repeating_header_rows_enabled",
        },
        30 => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.table_name_style",
        },
        31 => KnownField {
            kind: KnownFieldKind::Bool,
            wire: 0,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.style_apply_clears_all",
        },
        32 => KnownField {
            kind: KnownFieldKind::Bool,
            wire: 0,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.repeating_header_columns_enabled",
        },
        33 => KnownField {
            kind: KnownFieldKind::Fixed64,
            wire: 1,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.table_name_height",
        },
        34 => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.hidden_state_formula_owner_for_columns",
        },
        35 => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.hidden_state_formula_owner_for_rows",
        },
        36 => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.table_name_shape_style",
        },
        37 => KnownField {
            kind: KnownFieldKind::Bool,
            wire: 0,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.table_name_border_enabled",
        },
        38 => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.row_filter_set_pre_pivot",
        },
        39 => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.conditional_style_formula_owner_id",
        },
        40 => KnownField {
            kind: KnownFieldKind::Varint,
            wire: 0,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.number_of_filtered_rows",
        },
        41 => KnownField {
            kind: KnownFieldKind::Varint,
            wire: 0,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.number_of_user_hidden_rows",
        },
        42 => KnownField {
            kind: KnownFieldKind::Varint,
            wire: 0,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.number_of_user_hidden_columns",
        },
        43 => KnownField {
            kind: KnownFieldKind::String,
            wire: 2,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.from_table_id",
        },
        44 => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.sort_order",
        },
        45 => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.sort_rule_reference_tracker",
        },
        46 => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.base_column_row_uids",
        },
        47 => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.merge_owner",
        },
        48 => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.table_style_preset",
        },
        49 => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.stroke_sidecar",
        },
        50 => KnownField {
            kind: KnownFieldKind::Bool,
            wire: 0,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.was_cut",
        },
        51 => KnownField {
            kind: KnownFieldKind::Bool,
            wire: 0,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.preset_needs_strong_ownership",
        },
        52 => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.text_import_record",
        },
        60..=64 => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.category_style",
        },
        65..=69 => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.category_text_style",
        },
        70 => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.hidden_states_owner",
        },
        71..=75 => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.label_style",
        },
        76..=80 => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.label_text_style",
        },
        81 => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.category_owner_deprecated",
        },
        82 => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.pencil_annotation_owner",
        },
        83 => KnownField {
            kind: KnownFieldKind::String,
            wire: 2,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.from_group_by_uid",
        },
        84 => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.haunted_owner",
        },
        85 => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.pivot_owner",
        },
        86 => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.category_owner",
        },
        87 => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.pivot_body_summary_row_style",
        },
        88 => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.pivot_body_summary_column_style",
        },
        89 => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.pivot_header_column_summary_style",
        },
        90..=92 => KnownField {
            kind: KnownFieldKind::RepeatedVarint,
            wire: 0,
            required: false,
            repeated: true,
            name: "TST.TableModelArchive.repeated_uint32",
        },
        93 => KnownField {
            kind: KnownFieldKind::Bytes,
            wire: 2,
            required: false,
            repeated: false,
            name: "TST.TableModelArchive.spill_owner",
        },
        _ => return None,
    };
    Some(field)
}

/// Finite limits for one borrowed table-model discovery decode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_input_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    max_text_bytes: usize,
    recursion_limit: u32,
    max_output_bytes: usize,
    max_allocations: usize,
    max_retained_bytes: usize,
    max_scratch_bytes: usize,
}

impl DecodeOptions {
    /// Construct an explicit input/field/work/nesting policy.
    #[must_use]
    pub const fn new(
        max_input_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        recursion_limit: u32,
    ) -> Self {
        Self {
            max_input_bytes,
            max_fields,
            max_work_bytes,
            max_text_bytes: max_input_bytes,
            recursion_limit,
            max_output_bytes: doubled_at_least_one(max_input_bytes),
            max_allocations: 1,
            max_retained_bytes: doubled_at_least_one(max_input_bytes),
            max_scratch_bytes: doubled_at_least_one(max_input_bytes),
        }
    }

    /// Build a conservative source-sized policy for a trusted package scan.
    ///
    /// The work ceiling includes one strict source pass and the two private
    /// Buffa parity passes.  Callers with an aggregate operation budget should
    /// replace these values with their residual ceilings.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        Self::new(
            bytes,
            bytes,
            bytes
                .checked_mul(BUFFA_PARITY_PASSES + 1)
                .unwrap_or(usize::MAX)
                .max(1),
            MAX_RECURSION,
        )
    }

    /// Replace the input-byte ceiling.
    #[must_use]
    pub const fn with_max_input_bytes(mut self, maximum: usize) -> Self {
        self.max_input_bytes = maximum;
        self
    }

    /// Replace the field-record ceiling.
    #[must_use]
    pub const fn with_max_fields(mut self, maximum: usize) -> Self {
        self.max_fields = maximum;
        self
    }

    /// Replace the aggregate strict-plus-Buffa work ceiling.
    #[must_use]
    pub const fn with_max_work_bytes(mut self, maximum: usize) -> Self {
        self.max_work_bytes = maximum;
        self
    }

    /// Replace the aggregate UTF-8 text-byte ceiling for borrowed strings.
    #[must_use]
    pub const fn with_max_text_bytes(mut self, maximum: usize) -> Self {
        self.max_text_bytes = maximum;
        self
    }

    /// Replace the maximum unknown-group nesting depth.
    #[must_use]
    pub const fn with_recursion_limit(mut self, maximum: u32) -> Self {
        self.recursion_limit = maximum;
        self
    }

    /// Replace the candidate-output byte ceiling used by prepared rewrites.
    #[must_use]
    pub const fn with_max_output_bytes(mut self, maximum: usize) -> Self {
        self.max_output_bytes = maximum;
        self
    }

    /// Replace the logical output-allocation ceiling used by prepared rewrites.
    #[must_use]
    pub const fn with_max_allocations(mut self, maximum: usize) -> Self {
        self.max_allocations = maximum;
        self
    }

    /// Replace the retained-candidate byte ceiling used by prepared rewrites.
    #[must_use]
    pub const fn with_max_retained_bytes(mut self, maximum: usize) -> Self {
        self.max_retained_bytes = maximum;
        self
    }

    /// Replace the logical scratch-byte ceiling used by prepared rewrites.
    #[must_use]
    pub const fn with_max_scratch_bytes(mut self, maximum: usize) -> Self {
        self.max_scratch_bytes = maximum;
        self
    }

    /// Return the input-byte ceiling.
    #[must_use]
    pub const fn max_input_bytes(self) -> usize {
        self.max_input_bytes
    }

    /// Return the field-record ceiling.
    #[must_use]
    pub const fn max_fields(self) -> usize {
        self.max_fields
    }

    /// Return the aggregate work ceiling.
    #[must_use]
    pub const fn max_work_bytes(self) -> usize {
        self.max_work_bytes
    }

    /// Return the aggregate UTF-8 text-byte ceiling for borrowed strings.
    #[must_use]
    pub const fn max_text_bytes(self) -> usize {
        self.max_text_bytes
    }

    /// Return the nesting ceiling.
    #[must_use]
    pub const fn recursion_limit(self) -> u32 {
        self.recursion_limit
    }

    /// Return the candidate-output byte ceiling used by prepared rewrites.
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }

    /// Return the logical output-allocation ceiling used by prepared rewrites.
    #[must_use]
    pub const fn max_allocations(self) -> usize {
        self.max_allocations
    }

    /// Return the retained-candidate byte ceiling used by prepared rewrites.
    #[must_use]
    pub const fn max_retained_bytes(self) -> usize {
        self.max_retained_bytes
    }

    /// Return the logical scratch-byte ceiling used by prepared rewrites.
    #[must_use]
    pub const fn max_scratch_bytes(self) -> usize {
        self.max_scratch_bytes
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_input_bytes)
            // The strict scanner owns unknown-field accounting.  Buffa's
            // lazy view is only the generated parity check; use the same
            // finite field ceiling so accepted canonical unknown framing is
            // not rejected merely because the parity view sees it.
            .with_unknown_field_limit(self.max_fields)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }
}

/// Exact facts needed by table discovery; all strings borrow the source.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableModelDiscoverySnapshot<'source> {
    table_id: &'source str,
    table_name: &'source str,
    rows: u32,
    columns: u32,
}

impl<'source> TableModelDiscoverySnapshot<'source> {
    /// Required `TST.TableModelArchive.table_id`.
    #[must_use]
    pub const fn table_id(self) -> &'source str {
        self.table_id
    }

    /// Required `TST.TableModelArchive.table_name`.
    #[must_use]
    pub const fn table_name(self) -> &'source str {
        self.table_name
    }

    /// Required `TST.TableModelArchive.number_of_rows`.
    #[must_use]
    pub const fn rows(self) -> u32 {
        self.rows
    }

    /// Required `TST.TableModelArchive.number_of_columns`.
    #[must_use]
    pub const fn columns(self) -> u32 {
        self.columns
    }

    /// Compatibility spelling for callers that use the native field name.
    #[must_use]
    pub const fn number_of_rows(self) -> u32 {
        self.rows
    }

    /// Compatibility spelling for callers that use the native field name.
    #[must_use]
    pub const fn number_of_columns(self) -> u32 {
        self.columns
    }
}

/// Compatibility alias for package adapters that call these facts a model
/// snapshot.
pub type TableModelSnapshot<'source> = TableModelDiscoverySnapshot<'source>;

/// Compatibility alias for Keynote's private discovery terminology.
pub type TableModelFacts<'source> = TableModelDiscoverySnapshot<'source>;

/// Exact successful accounting for one source projection.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct DecodeReport {
    input_bytes: usize,
    fields: usize,
    work_bytes: usize,
    text_bytes: usize,
    max_depth: u32,
}

impl DecodeReport {
    /// Bytes borrowed from the caller's source.
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    /// Field records visited, including fields nested in unknown groups.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Aggregate strict-scan and Buffa parity work reservation.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// UTF-8 bytes borrowed for the model's string fields.
    #[must_use]
    pub const fn text_bytes(self) -> usize {
        self.text_bytes
    }

    /// Greatest unknown-group depth reached.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    /// This projection never allocates an output container.
    #[must_use]
    pub const fn allocations(self) -> usize {
        0
    }

    /// The snapshot retains only source borrows, not owned bytes.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        0
    }

    /// The strict scanner uses caller-owned slices and fixed stack state.
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        0
    }
}

/// Typed resource observation from a bounded discovery decode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    /// The source or configured Buffa message exceeded the byte ceiling.
    Bytes { observed: usize, maximum: usize },
    /// The strict scanner visited too many field records.
    Fields { observed: usize, maximum: usize },
    /// The strict and parity passes exceeded the work ceiling.
    Work { observed: usize, maximum: usize },
    /// Borrowed UTF-8 text exceeded the text-byte ceiling.
    Text { observed: usize, maximum: usize },
    /// The prepared rewrite exceeded its candidate-output byte ceiling.
    Output { observed: usize, maximum: usize },
    /// The prepared rewrite exceeded its logical allocation ceiling.
    Allocations { observed: usize, maximum: usize },
    /// The prepared rewrite exceeded its retained-byte ceiling.
    Retained { observed: usize, maximum: usize },
    /// The prepared rewrite exceeded its scratch-byte ceiling.
    Scratch { observed: usize, maximum: usize },
    /// An unknown group exceeded the configured nesting ceiling.
    Nesting { observed: u32, maximum: u32 },
}

/// Strict discovery failure.  Diagnostics contain no IDs or source bytes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Wire(&'static str),
    Limit(DecodeLimit),
    Missing(&'static str),
    Duplicate(&'static str),
    NonCanonical(&'static str),
    FingerprintMismatch,
    Projection,
}

impl DecodeError {
    const fn wire(message: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::Wire(message),
        }
    }

    const fn limit(limit: DecodeLimit) -> Self {
        Self {
            kind: DecodeErrorKind::Limit(limit),
        }
    }

    const fn missing(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::Missing(field),
        }
    }

    const fn duplicate(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::Duplicate(field),
        }
    }

    const fn noncanonical(reason: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::NonCanonical(reason),
        }
    }

    const fn fingerprint_mismatch() -> Self {
        Self {
            kind: DecodeErrorKind::FingerprintMismatch,
        }
    }

    const fn projection() -> Self {
        Self {
            kind: DecodeErrorKind::Projection,
        }
    }

    /// Return the typed resource failure, if this decode was bounded out.
    #[must_use]
    pub const fn resource_limit(&self) -> Option<DecodeLimit> {
        if let DecodeErrorKind::Limit(limit) = self.kind {
            Some(limit)
        } else {
            None
        }
    }

    /// Return the missing required field, if applicable.
    #[must_use]
    pub const fn missing_required_field(&self) -> Option<&'static str> {
        if let DecodeErrorKind::Missing(field) = self.kind {
            Some(field)
        } else {
            None
        }
    }

    /// Return the duplicated singular field, if applicable.
    #[must_use]
    pub const fn duplicate_singular_field(&self) -> Option<&'static str> {
        if let DecodeErrorKind::Duplicate(field) = self.kind {
            Some(field)
        } else {
            None
        }
    }

    /// Return the stable canonicality reason, if applicable.
    #[must_use]
    pub const fn noncanonical_reason(&self) -> Option<&'static str> {
        if let DecodeErrorKind::NonCanonical(reason) = self.kind {
            Some(reason)
        } else {
            None
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self.kind {
            DecodeErrorKind::Wire(message) => formatter.write_str(message),
            DecodeErrorKind::Limit(DecodeLimit::Bytes { observed, maximum }) => write!(
                formatter,
                "table-model discovery input uses {observed} bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Fields { observed, maximum }) => write!(
                formatter,
                "table-model discovery visited {observed} fields; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Work { observed, maximum }) => write!(
                formatter,
                "table-model discovery requires {observed} work bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Text { observed, maximum }) => write!(
                formatter,
                "table-model discovery borrows {observed} text bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Output { observed, maximum }) => write!(
                formatter,
                "table-model name rewrite emits {observed} bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Allocations { observed, maximum }) => write!(
                formatter,
                "table-model name rewrite requires {observed} allocations; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Retained { observed, maximum }) => write!(
                formatter,
                "table-model name rewrite retains {observed} bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Scratch { observed, maximum }) => write!(
                formatter,
                "table-model name rewrite requires {observed} scratch bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Nesting { observed, maximum }) => write!(
                formatter,
                "table-model discovery reached nesting {observed}; maximum is {maximum}"
            ),
            DecodeErrorKind::Missing(field) => write!(formatter, "missing required field {field}"),
            DecodeErrorKind::Duplicate(field) => {
                write!(formatter, "duplicate singular field {field}")
            },
            DecodeErrorKind::NonCanonical(reason) => {
                write!(formatter, "non-canonical protobuf representation: {reason}")
            },
            DecodeErrorKind::FingerprintMismatch => {
                formatter.write_str("table-model rewrite source fingerprint changed")
            },
            DecodeErrorKind::Projection => {
                formatter.write_str("table-model discovery projection disagrees with Buffa")
            },
        }
    }
}

impl std::error::Error for DecodeError {}

/// Borrowed table-name rewrite request.  The requested name remains owned by
/// the caller until the prepared operation is executed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableModelNameWrite<'name> {
    name: &'name str,
    expected_fingerprint: Option<u64>,
}

impl<'name> TableModelNameWrite<'name> {
    /// Build a name rewrite request without changing the source fingerprint
    /// policy.
    #[must_use]
    pub const fn new(name: &'name str) -> Self {
        Self {
            name,
            expected_fingerprint: None,
        }
    }

    /// Require the complete source payload to have this fingerprint before a
    /// rewrite is prepared.
    #[must_use]
    pub const fn with_fingerprint(mut self, expected_fingerprint: u64) -> Self {
        self.expected_fingerprint = Some(expected_fingerprint);
        self
    }

    /// Return the requested table name.
    #[must_use]
    pub const fn name(self) -> &'name str {
        self.name
    }

    /// Return the optional source-fingerprint guard.
    #[must_use]
    pub const fn expected_fingerprint(self) -> Option<u64> {
        self.expected_fingerprint
    }
}

impl<'name> From<&'name str> for TableModelNameWrite<'name> {
    fn from(name: &'name str) -> Self {
        Self::new(name)
    }
}

/// Aggregate accounting for one prepared table-name rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableModelNameRewriteReport {
    input_bytes: usize,
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    text_bytes: usize,
    max_depth: u32,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
    changed: bool,
    source_fingerprint: u64,
}

impl TableModelNameRewriteReport {
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
    pub const fn text_bytes(self) -> usize {
        self.text_bytes
    }

    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
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

    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    #[must_use]
    pub const fn source_fingerprint(self) -> u64 {
        self.source_fingerprint
    }
}

/// Exact ceilings required by one prepared table-name rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableModelNameRewriteRequirements {
    input_bytes: usize,
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    text_bytes: usize,
    max_depth: u32,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

impl TableModelNameRewriteRequirements {
    /// Build execution limits that accept exactly these requirements.
    #[must_use]
    pub const fn exact(self) -> TableModelNameRewriteLimits {
        TableModelNameRewriteLimits {
            input_bytes: self.input_bytes,
            output_bytes: self.output_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            text_bytes: self.text_bytes,
            max_depth: self.max_depth,
            allocations: self.allocations,
            retained_bytes: self.retained_bytes,
            scratch_bytes: self.scratch_bytes,
        }
    }

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
    pub const fn text_bytes(self) -> usize {
        self.text_bytes
    }

    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
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

/// Caller-provided ceilings replayed before the rewrite output is allocated.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableModelNameRewriteLimits {
    input_bytes: usize,
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    text_bytes: usize,
    max_depth: u32,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

impl TableModelNameRewriteLimits {
    /// Start with no additional execution restriction.
    #[must_use]
    pub const fn unrestricted() -> Self {
        Self {
            input_bytes: usize::MAX,
            output_bytes: usize::MAX,
            fields: usize::MAX,
            work_bytes: usize::MAX,
            text_bytes: usize::MAX,
            max_depth: u32::MAX,
            allocations: usize::MAX,
            retained_bytes: usize::MAX,
            scratch_bytes: usize::MAX,
        }
    }

    #[must_use]
    pub const fn with_input_bytes(mut self, maximum: usize) -> Self {
        self.input_bytes = maximum;
        self
    }

    #[must_use]
    pub const fn with_output_bytes(mut self, maximum: usize) -> Self {
        self.output_bytes = maximum;
        self
    }

    #[must_use]
    pub const fn with_fields(mut self, maximum: usize) -> Self {
        self.fields = maximum;
        self
    }

    #[must_use]
    pub const fn with_work_bytes(mut self, maximum: usize) -> Self {
        self.work_bytes = maximum;
        self
    }

    #[must_use]
    pub const fn with_text_bytes(mut self, maximum: usize) -> Self {
        self.text_bytes = maximum;
        self
    }

    #[must_use]
    pub const fn with_max_depth(mut self, maximum: u32) -> Self {
        self.max_depth = maximum;
        self
    }

    #[must_use]
    pub const fn with_allocations(mut self, maximum: usize) -> Self {
        self.allocations = maximum;
        self
    }

    #[must_use]
    pub const fn with_retained_bytes(mut self, maximum: usize) -> Self {
        self.retained_bytes = maximum;
        self
    }

    #[must_use]
    pub const fn with_scratch_bytes(mut self, maximum: usize) -> Self {
        self.scratch_bytes = maximum;
        self
    }
}

/// Output bytes and accounting produced by one prepared execution.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableModelNameRewriteOutput {
    bytes: Vec<u8>,
    report: TableModelNameRewriteReport,
}

impl TableModelNameRewriteOutput {
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }

    #[must_use]
    pub const fn report(&self) -> TableModelNameRewriteReport {
        self.report
    }
}

/// Strictly prepared raw-preserving table-name rewrite.
#[derive(Debug, Clone, Copy)]
pub struct PreparedTableModelNameRewrite<'source, 'name> {
    source: &'source [u8],
    name: TableModelNameWrite<'name>,
    name_field_start: usize,
    name_field_end: usize,
    options: DecodeOptions,
    report: TableModelNameRewriteReport,
    requirements: TableModelNameRewriteRequirements,
    current: TableModelDiscoverySnapshot<'source>,
}

impl<'source, 'name> PreparedTableModelNameRewrite<'source, 'name> {
    /// Return preparation and execution accounting.
    #[must_use]
    pub const fn prepare_report(&self) -> TableModelNameRewriteReport {
        self.report
    }

    /// Return exact ceilings consumed by execution.
    #[must_use]
    pub const fn execution_requirements(&self) -> TableModelNameRewriteRequirements {
        self.requirements
    }

    /// Return the source fingerprint captured during preparation.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.report.source_fingerprint
    }

    /// Execute the prepared rewrite after checking every advertised ceiling.
    pub fn execute(
        self,
        limits: TableModelNameRewriteLimits,
    ) -> Result<TableModelNameRewriteOutput, DecodeError> {
        check_name_rewrite_limits(self.requirements, limits)?;
        if table_model_source_fingerprint(self.source) != self.report.source_fingerprint {
            return Err(DecodeError::fingerprint_mismatch());
        }

        let mut output = Vec::new();
        output
            .try_reserve_exact(self.requirements.output_bytes)
            .map_err(|_error| DecodeError::wire("table-model name rewrite allocation failed"))?;
        if self.report.changed {
            output.extend_from_slice(&self.source[..self.name_field_start]);
            push_varint((u64::from(TABLE_NAME_FIELD) << 3) | 2, &mut output);
            push_varint(
                u64::try_from(self.name.name.len()).map_err(|_error| DecodeError::projection())?,
                &mut output,
            );
            output.extend_from_slice(self.name.name.as_bytes());
            output.extend_from_slice(&self.source[self.name_field_end..]);
        } else {
            output.extend_from_slice(self.source);
        }
        if output.len() != self.requirements.output_bytes {
            return Err(DecodeError::projection());
        }

        let readback_options = self
            .options
            .with_max_input_bytes(output.len())
            .with_max_fields(self.requirements.fields)
            .with_max_work_bytes(self.requirements.work_bytes)
            .with_max_text_bytes(self.requirements.text_bytes)
            .with_max_output_bytes(output.len())
            .with_max_allocations(self.requirements.allocations)
            .with_max_retained_bytes(self.requirements.retained_bytes)
            .with_max_scratch_bytes(self.requirements.scratch_bytes);
        let (readback, _report) = decode_table_model_with_report(&output, readback_options)?;
        if readback.table_id() != self.current.table_id()
            || readback.rows() != self.current.rows()
            || readback.columns() != self.current.columns()
            || readback.table_name() != self.name.name
        {
            return Err(DecodeError::projection());
        }

        Ok(TableModelNameRewriteOutput {
            bytes: output,
            report: self.report,
        })
    }
}

/// Prepare a strict raw-preserving rewrite of required field 8
/// (`TST.TableModelArchive.table_name`).
pub fn prepare_table_model_name_rewrite<'source, 'name>(
    source: &'source [u8],
    write: impl Into<TableModelNameWrite<'name>>,
    options: DecodeOptions,
) -> Result<PreparedTableModelNameRewrite<'source, 'name>, DecodeError> {
    let write = write.into();
    let (parsed, source_report) = decode_parsed_model_with_report(source, options)?;
    if write
        .expected_fingerprint()
        .is_some_and(|expected| expected != table_model_source_fingerprint(source))
    {
        return Err(DecodeError::fingerprint_mismatch());
    }

    let changed = parsed.table_name != write.name();
    let old_field_len = parsed
        .name_field_end
        .checked_sub(parsed.name_field_start)
        .ok_or_else(DecodeError::projection)?;
    let new_field_len = encoded_length_delimited_field_len(TABLE_NAME_FIELD, write.name().len())?;
    let output_bytes = source
        .len()
        .checked_sub(old_field_len)
        .and_then(|length| length.checked_add(new_field_len))
        .ok_or_else(DecodeError::projection)?;
    let candidate_text_bytes = source_report
        .text_bytes
        .checked_sub(parsed.table_name.len())
        .and_then(|text| text.checked_add(write.name().len()))
        .ok_or_else(DecodeError::projection)?;
    if candidate_text_bytes > options.max_text_bytes {
        return Err(DecodeError::limit(DecodeLimit::Text {
            observed: candidate_text_bytes,
            maximum: options.max_text_bytes,
        }));
    }

    let candidate_work = output_bytes
        .checked_mul(BUFFA_PARITY_PASSES + 1)
        .ok_or_else(DecodeError::projection)?;
    let rewrite_work = output_bytes;
    let fields = source_report
        .fields
        .checked_mul(2)
        .ok_or_else(DecodeError::projection)?;
    let work_bytes = source_report
        .work_bytes
        .checked_add(candidate_work)
        .and_then(|work| work.checked_add(rewrite_work))
        .ok_or_else(DecodeError::projection)?;
    let text_bytes = source_report
        .text_bytes
        .checked_add(candidate_text_bytes)
        .ok_or_else(DecodeError::projection)?;
    let scratch_bytes = source
        .len()
        .checked_add(output_bytes)
        .ok_or_else(DecodeError::projection)?;
    let requirements = TableModelNameRewriteRequirements {
        input_bytes: source.len(),
        output_bytes,
        fields,
        work_bytes,
        text_bytes,
        max_depth: source_report.max_depth,
        allocations: 1,
        retained_bytes: output_bytes,
        scratch_bytes,
    };
    check_name_rewrite_limits(
        requirements,
        TableModelNameRewriteLimits {
            input_bytes: options.max_input_bytes,
            output_bytes: options.max_output_bytes,
            fields: options.max_fields,
            work_bytes: options.max_work_bytes,
            text_bytes: options.max_text_bytes,
            max_depth: options.recursion_limit,
            allocations: options.max_allocations,
            retained_bytes: options.max_retained_bytes,
            scratch_bytes: options.max_scratch_bytes,
        },
    )?;
    let report = TableModelNameRewriteReport {
        input_bytes: source.len(),
        output_bytes,
        fields,
        work_bytes,
        text_bytes,
        max_depth: source_report.max_depth,
        allocations: 1,
        retained_bytes: output_bytes,
        scratch_bytes,
        changed,
        source_fingerprint: table_model_source_fingerprint(source),
    };
    Ok(PreparedTableModelNameRewrite {
        source,
        name: write,
        name_field_start: parsed.name_field_start,
        name_field_end: parsed.name_field_end,
        options,
        report,
        requirements,
        current: TableModelDiscoverySnapshot {
            table_id: parsed.table_id,
            table_name: parsed.table_name,
            rows: parsed.rows,
            columns: parsed.columns,
        },
    })
}

/// Prepare and execute a strict table-name rewrite in one call.
pub fn rewrite_table_model_name(
    source: &[u8],
    name: &str,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    let prepared =
        prepare_table_model_name_rewrite(source, TableModelNameWrite::new(name), options)?;
    Ok(prepared
        .execute(prepared.execution_requirements().exact())?
        .into_bytes())
}

/// Prepare and execute a strict table-name rewrite with accounting.
pub fn rewrite_table_model_name_with_report(
    source: &[u8],
    name: &str,
    options: DecodeOptions,
) -> Result<(Vec<u8>, TableModelNameRewriteReport), DecodeError> {
    let prepared =
        prepare_table_model_name_rewrite(source, TableModelNameWrite::new(name), options)?;
    let output = prepared.execute(prepared.execution_requirements().exact())?;
    let report = output.report();
    Ok((output.into_bytes(), report))
}

/// Stable allocation-free fingerprint for a complete table-model payload.
#[must_use]
pub fn table_model_source_fingerprint(source: &[u8]) -> u64 {
    let mut fingerprint = 14_695_981_039_346_656_037u64;
    for byte in source {
        fingerprint ^= u64::from(*byte);
        fingerprint = fingerprint.wrapping_mul(1_099_511_628_211u64);
    }
    fingerprint
}

/// Decode the strict borrowed table-model discovery facts.
pub fn decode_table_model(
    source: &[u8],
    options: DecodeOptions,
) -> Result<TableModelSnapshot<'_>, DecodeError> {
    Ok(decode_table_model_with_report(source, options)?.0)
}

/// Decode the strict borrowed facts and aggregate traversal accounting.
pub fn decode_table_model_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(TableModelSnapshot<'_>, DecodeReport), DecodeError> {
    let (parsed, report) = decode_parsed_model_with_report(source, options)?;
    Ok((
        TableModelDiscoverySnapshot {
            table_id: parsed.table_id,
            table_name: parsed.table_name,
            rows: parsed.rows,
            columns: parsed.columns,
        },
        report,
    ))
}

fn decode_parsed_model_with_report<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<(ParsedModel<'source>, DecodeReport), DecodeError> {
    validate_input(source, options)?;
    let mut budget = Budget::new(options, source.len());
    let parsed = parse_model(source, &mut budget, 0)?;
    let work_bytes = budget.finish()?;

    // Force both existing lazy Buffa projections only after the complete
    // strict pass.  They are generated private implementation details; the
    // source scan remains the preservation and unknown-field authority.
    let names: names_projection::NumbersTableModelArchiveLazyView<'_> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::projection())?;
    let dimensions: dimensions_projection::NumbersTableHeaderSettingsArchiveLazyView<'_> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::projection())?;
    if !names.has_table_id()
        || !names.has_table_name()
        || !dimensions.has_number_of_rows()
        || !dimensions.has_number_of_columns()
        || names.table_id != parsed.table_id
        || names.table_name != parsed.table_name
        || dimensions.number_of_rows != parsed.rows
        || dimensions.number_of_columns != parsed.columns
    {
        return Err(DecodeError::projection());
    }

    Ok((
        parsed,
        DecodeReport {
            input_bytes: source.len(),
            fields: budget.fields,
            work_bytes,
            text_bytes: budget.text_bytes,
            max_depth: budget.max_depth,
        },
    ))
}

/// Explicit discovery-facts spelling for Keynote adapters.
pub fn decode_table_model_facts(
    source: &[u8],
    options: DecodeOptions,
) -> Result<TableModelFacts<'_>, DecodeError> {
    decode_table_model(source, options)
}

/// Explicit discovery-facts spelling with traversal accounting.
pub fn decode_table_model_facts_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(TableModelFacts<'_>, DecodeReport), DecodeError> {
    decode_table_model_with_report(source, options)
}

/// Explicit table-model-discovery spelling for package adapters.
pub fn decode_table_model_discovery(
    source: &[u8],
    options: DecodeOptions,
) -> Result<TableModelDiscoverySnapshot<'_>, DecodeError> {
    decode_table_model(source, options)
}

/// Explicit table-model-discovery spelling with traversal accounting.
pub fn decode_table_model_discovery_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(TableModelDiscoverySnapshot<'_>, DecodeReport), DecodeError> {
    decode_table_model_with_report(source, options)
}

fn validate_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    let hard_maximum =
        usize::try_from(buffa::MAX_MESSAGE_BYTES).map_err(|_error| DecodeError::projection())?;
    if options.max_input_bytes > hard_maximum || source.len() > options.max_input_bytes {
        return Err(DecodeError::limit(DecodeLimit::Bytes {
            observed: source.len(),
            maximum: options.max_input_bytes.min(hard_maximum),
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

#[derive(Debug)]
struct Budget {
    fields: usize,
    work_bytes: usize,
    text_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    max_text_bytes: usize,
    max_depth: u32,
    recursion_limit: u32,
    source_len: usize,
    source_cursor: usize,
}

impl Budget {
    const fn new(options: DecodeOptions, source_len: usize) -> Self {
        Self {
            fields: 0,
            work_bytes: 0,
            text_bytes: 0,
            max_fields: options.max_fields,
            max_work_bytes: options.max_work_bytes,
            max_text_bytes: options.max_text_bytes,
            max_depth: 0,
            recursion_limit: options.recursion_limit,
            source_len,
            source_cursor: 0,
        }
    }

    fn field(&mut self) -> Result<(), DecodeError> {
        let observed = self.fields.checked_add(1).ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Fields {
                observed: usize::MAX,
                maximum: self.max_fields,
            })
        })?;
        if observed > self.max_fields {
            return Err(DecodeError::limit(DecodeLimit::Fields {
                observed,
                maximum: self.max_fields,
            }));
        }
        self.fields = observed;
        Ok(())
    }

    fn source_pass(&mut self, bytes: usize) -> Result<(), DecodeError> {
        let observed = self.work_bytes.checked_add(bytes).ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Work {
                observed: usize::MAX,
                maximum: self.max_work_bytes,
            })
        })?;
        if observed > self.max_work_bytes {
            return Err(DecodeError::limit(DecodeLimit::Work {
                observed,
                maximum: self.max_work_bytes,
            }));
        }
        self.work_bytes = observed;
        Ok(())
    }

    fn source_progress(&mut self, remaining: &[u8]) -> Result<(), DecodeError> {
        let consumed = self
            .source_len
            .checked_sub(remaining.len())
            .ok_or_else(|| DecodeError::wire("source cursor underflow"))?;
        let delta = consumed
            .checked_sub(self.source_cursor)
            .ok_or_else(|| DecodeError::wire("source cursor moved backwards"))?;
        self.source_cursor = consumed;
        self.source_pass(delta)
    }

    fn text(&mut self, bytes: usize) -> Result<(), DecodeError> {
        let observed = self.text_bytes.checked_add(bytes).ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Text {
                observed: usize::MAX,
                maximum: self.max_text_bytes,
            })
        })?;
        if observed > self.max_text_bytes {
            return Err(DecodeError::limit(DecodeLimit::Text {
                observed,
                maximum: self.max_text_bytes,
            }));
        }
        self.text_bytes = observed;
        Ok(())
    }

    fn depth(&mut self, depth: u32) -> Result<(), DecodeError> {
        self.max_depth = self.max_depth.max(depth);
        if depth > self.recursion_limit {
            return Err(DecodeError::limit(DecodeLimit::Nesting {
                observed: depth,
                maximum: self.recursion_limit,
            }));
        }
        Ok(())
    }

    fn finish(&self) -> Result<usize, DecodeError> {
        let parity_work = self
            .source_len
            .checked_mul(BUFFA_PARITY_PASSES)
            .and_then(|bytes| self.work_bytes.checked_add(bytes))
            .ok_or_else(|| {
                DecodeError::limit(DecodeLimit::Work {
                    observed: usize::MAX,
                    maximum: self.max_work_bytes,
                })
            })?;
        if parity_work > self.max_work_bytes {
            return Err(DecodeError::limit(DecodeLimit::Work {
                observed: parity_work,
                maximum: self.max_work_bytes,
            }));
        }
        Ok(parity_work)
    }
}

#[derive(Debug, Clone, Copy)]
struct ParsedModel<'source> {
    table_id: &'source str,
    table_name: &'source str,
    rows: u32,
    columns: u32,
    name_field_start: usize,
    name_field_end: usize,
}

fn parse_model<'source>(
    source: &'source [u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<ParsedModel<'source>, DecodeError> {
    budget.depth(depth)?;
    let mut remaining = source;
    let mut table_id = None;
    let mut table_name = None;
    let mut rows = None;
    let mut columns = None;
    let mut name_field_start = None;
    let mut name_field_end = None;
    let mut required_fields = 0u32;
    let mut seen_known_fields = 0u128;
    while !remaining.is_empty() {
        let field_start = source
            .len()
            .checked_sub(remaining.len())
            .ok_or_else(|| DecodeError::wire("source cursor underflow"))?;
        let (tag, key_canonical) = read_varint(&mut remaining)?;
        if !key_canonical {
            return Err(DecodeError::noncanonical("protobuf field key"));
        }
        let raw_tag =
            u32::try_from(tag).map_err(|_error| DecodeError::wire("invalid field number"))?;
        let number = raw_tag >> 3;
        let wire = raw_tag & 7;
        if number == 0 || number > MAX_FIELD_NUMBER {
            return Err(DecodeError::wire("invalid field number"));
        }
        budget.field()?;
        if let Some(known) = known_field(number) {
            require_wire(number, wire, known.wire)?;
            if !known.repeated {
                mark_known(&mut seen_known_fields, number, known.name)?;
            }
            if known.required {
                mark_required(&mut required_fields, number, known.name)?;
            }
            match known.kind {
                KnownFieldKind::String => {
                    let value = read_string(&mut remaining, budget)?;
                    match number {
                        TABLE_ID_FIELD => table_id = Some(value),
                        TABLE_NAME_FIELD => {
                            table_name = Some(value);
                            name_field_start = Some(field_start);
                            name_field_end = Some(
                                source
                                    .len()
                                    .checked_sub(remaining.len())
                                    .ok_or_else(|| DecodeError::wire("source cursor underflow"))?,
                            );
                        },
                        _ => {},
                    }
                },
                KnownFieldKind::Varint | KnownFieldKind::RepeatedVarint => {
                    let value = read_u32(&mut remaining)?;
                    match number {
                        TABLE_ROWS_FIELD => rows = Some(value),
                        TABLE_COLUMNS_FIELD => columns = Some(value),
                        _ => {},
                    }
                },
                KnownFieldKind::Bool => {
                    let _ = read_bool(&mut remaining)?;
                },
                KnownFieldKind::Fixed64 => take(&mut remaining, 8)?,
                KnownFieldKind::Bytes => {
                    let _ = read_length_delimited(&mut remaining)?;
                },
            }
        } else {
            skip_unknown(&mut remaining, number, wire, depth, budget)?;
        }
        budget.source_progress(remaining)?;
    }
    if required_fields & REQUIRED_FIELDS_MASK != REQUIRED_FIELDS_MASK {
        let missing = [
            TABLE_ID_FIELD,
            TABLE_STYLE_FIELD,
            BASE_DATA_STORE_FIELD,
            TABLE_ROWS_FIELD,
            TABLE_COLUMNS_FIELD,
            TABLE_NAME_FIELD,
            DEFAULT_ROW_HEIGHT_FIELD,
            DEFAULT_COLUMN_WIDTH_FIELD,
            BODY_CELL_STYLE_FIELD,
            HEADER_ROW_STYLE_FIELD,
            HEADER_COLUMN_STYLE_FIELD,
            FOOTER_ROW_STYLE_FIELD,
            BODY_TEXT_STYLE_FIELD,
            HEADER_ROW_TEXT_STYLE_FIELD,
            HEADER_COLUMN_TEXT_STYLE_FIELD,
            FOOTER_ROW_TEXT_STYLE_FIELD,
        ]
        .into_iter()
        .find(|field| {
            required_fields
                & 1u32
                    .checked_shl(*field)
                    .expect("TableModel required fields fit the presence mask")
                == 0
        });
        if let Some(field) = missing {
            return Err(DecodeError::missing(required_field_name(field)));
        }
    }
    Ok(ParsedModel {
        table_id: table_id.ok_or_else(|| DecodeError::missing("TST.TableModelArchive.table_id"))?,
        table_name: table_name
            .ok_or_else(|| DecodeError::missing("TST.TableModelArchive.table_name"))?,
        rows: rows.ok_or_else(|| DecodeError::missing("TST.TableModelArchive.number_of_rows"))?,
        columns: columns
            .ok_or_else(|| DecodeError::missing("TST.TableModelArchive.number_of_columns"))?,
        name_field_start: name_field_start
            .ok_or_else(|| DecodeError::missing("TST.TableModelArchive.table_name"))?,
        name_field_end: name_field_end
            .ok_or_else(|| DecodeError::missing("TST.TableModelArchive.table_name"))?,
    })
}

fn mark_required(seen: &mut u32, field: u32, name: &'static str) -> Result<(), DecodeError> {
    let bit = 1u32
        .checked_shl(field)
        .ok_or_else(|| DecodeError::wire("required field number is out of range"))?;
    if *seen & bit != 0 {
        return Err(DecodeError::duplicate(name));
    }
    *seen |= bit;
    Ok(())
}

fn mark_known(seen: &mut u128, field: u32, name: &'static str) -> Result<(), DecodeError> {
    let bit = 1u128
        .checked_shl(field)
        .ok_or_else(|| DecodeError::wire("known field number is out of range"))?;
    if *seen & bit != 0 {
        return Err(DecodeError::duplicate(name));
    }
    *seen |= bit;
    Ok(())
}

const fn required_field_name(field: u32) -> &'static str {
    match field {
        TABLE_ID_FIELD => "TST.TableModelArchive.table_id",
        TABLE_STYLE_FIELD => "TST.TableModelArchive.table_style",
        BASE_DATA_STORE_FIELD => "TST.TableModelArchive.base_data_store",
        TABLE_ROWS_FIELD => "TST.TableModelArchive.number_of_rows",
        TABLE_COLUMNS_FIELD => "TST.TableModelArchive.number_of_columns",
        TABLE_NAME_FIELD => "TST.TableModelArchive.table_name",
        DEFAULT_ROW_HEIGHT_FIELD => "TST.TableModelArchive.default_row_height",
        DEFAULT_COLUMN_WIDTH_FIELD => "TST.TableModelArchive.default_column_width",
        BODY_CELL_STYLE_FIELD => "TST.TableModelArchive.body_cell_style",
        HEADER_ROW_STYLE_FIELD => "TST.TableModelArchive.header_row_style",
        HEADER_COLUMN_STYLE_FIELD => "TST.TableModelArchive.header_column_style",
        FOOTER_ROW_STYLE_FIELD => "TST.TableModelArchive.footer_row_style",
        BODY_TEXT_STYLE_FIELD => "TST.TableModelArchive.body_text_style",
        HEADER_ROW_TEXT_STYLE_FIELD => "TST.TableModelArchive.header_row_text_style",
        HEADER_COLUMN_TEXT_STYLE_FIELD => "TST.TableModelArchive.header_column_text_style",
        FOOTER_ROW_TEXT_STYLE_FIELD => "TST.TableModelArchive.footer_row_text_style",
        _ => "TST.TableModelArchive.required_field",
    }
}

fn require_wire(number: u32, actual: u32, expected: u32) -> Result<(), DecodeError> {
    if actual == expected {
        Ok(())
    } else {
        Err(DecodeError::wire(match (number, expected) {
            (TABLE_ID_FIELD, 2) => "table_id has the wrong wire type",
            (TABLE_ROWS_FIELD, 0) => "number_of_rows has the wrong wire type",
            (TABLE_COLUMNS_FIELD, 0) => "number_of_columns has the wrong wire type",
            (TABLE_NAME_FIELD, 2) => "table_name has the wrong wire type",
            _ => "known field has the wrong wire type",
        }))
    }
}

fn read_string<'source>(
    source: &mut &'source [u8],
    budget: &mut Budget,
) -> Result<&'source str, DecodeError> {
    let bytes = read_length_delimited(source)?;
    budget.text(bytes.len())?;
    str::from_utf8(bytes).map_err(|_error| DecodeError::noncanonical("string is not valid UTF-8"))
}

fn read_u32(source: &mut &[u8]) -> Result<u32, DecodeError> {
    let (value, canonical) = read_varint(source)?;
    if !canonical {
        return Err(DecodeError::noncanonical("protobuf varint value"));
    }
    u32::try_from(value).map_err(|_error| DecodeError::noncanonical("uint32 scalar exceeds u32"))
}

fn read_bool(source: &mut &[u8]) -> Result<bool, DecodeError> {
    match read_u32(source)? {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(DecodeError::noncanonical("bool scalar is not zero or one")),
    }
}

fn read_length_delimited<'source>(
    source: &mut &'source [u8],
) -> Result<&'source [u8], DecodeError> {
    let (length, canonical) = read_varint(source)?;
    if !canonical {
        return Err(DecodeError::noncanonical("length-delimited size"));
    }
    let length =
        usize::try_from(length).map_err(|_error| DecodeError::wire("message too large"))?;
    if source.len() < length {
        return Err(DecodeError::wire("unexpected end of message"));
    }
    let (value, rest) = source.split_at(length);
    *source = rest;
    Ok(value)
}

fn skip_unknown(
    source: &mut &[u8],
    field_number: u32,
    wire: u32,
    depth: u32,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    match wire {
        0 => {
            let (_, canonical) = read_varint(source)?;
            if !canonical {
                return Err(DecodeError::noncanonical("protobuf varint value"));
            }
            Ok(())
        },
        1 => take(source, 8),
        2 => {
            let _ = read_length_delimited(source)?;
            Ok(())
        },
        3 => {
            let nested_depth = depth.checked_add(1).ok_or_else(|| {
                DecodeError::limit(DecodeLimit::Nesting {
                    observed: u32::MAX,
                    maximum: budget.recursion_limit,
                })
            })?;
            budget.depth(nested_depth)?;
            loop {
                if source.is_empty() {
                    return Err(DecodeError::wire("unexpected end of unknown group"));
                }
                let (tag, canonical) = read_varint(source)?;
                if !canonical {
                    return Err(DecodeError::noncanonical("protobuf field key"));
                }
                let raw_tag = u32::try_from(tag)
                    .map_err(|_error| DecodeError::wire("invalid field number"))?;
                let number = raw_tag >> 3;
                let nested_wire = raw_tag & 7;
                if number == 0 || number > MAX_FIELD_NUMBER {
                    return Err(DecodeError::wire("invalid field number"));
                }
                budget.field()?;
                if nested_wire == 4 {
                    if number == field_number {
                        budget.source_progress(source)?;
                        return Ok(());
                    }
                    return Err(DecodeError::wire("mismatched unknown group end"));
                }
                skip_unknown(source, number, nested_wire, nested_depth, budget)?;
                budget.source_progress(source)?;
            }
        },
        4 => Err(DecodeError::wire("unexpected unknown group end")),
        5 => take(source, 4),
        _ => Err(DecodeError::wire("invalid wire type")),
    }
}

fn take(source: &mut &[u8], length: usize) -> Result<(), DecodeError> {
    if source.len() < length {
        return Err(DecodeError::wire("unexpected end of message"));
    }
    *source = &source[length..];
    Ok(())
}

fn read_varint(source: &mut &[u8]) -> Result<(u64, bool), DecodeError> {
    let original = *source;
    let mut value = 0u64;
    for index in 0..10 {
        let byte = *original
            .get(index)
            .ok_or_else(|| DecodeError::wire("unexpected end of varint"))?;
        if index == 9 && byte > 1 {
            return Err(DecodeError::wire("varint is too long"));
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            *source = &original[index + 1..];
            let mut remaining = value;
            let mut canonical_length = 1usize;
            while remaining >= 128 {
                remaining >>= 7;
                canonical_length += 1;
            }
            return Ok((value, canonical_length == index + 1));
        }
    }
    Err(DecodeError::wire("varint is too long"))
}

fn encoded_length_delimited_field_len(
    field_number: u32,
    payload_len: usize,
) -> Result<usize, DecodeError> {
    let key = (u64::from(field_number) << 3) | 2;
    varint_len(key)
        .checked_add(varint_len(
            u64::try_from(payload_len).map_err(|_error| DecodeError::projection())?,
        ))
        .and_then(|length| length.checked_add(payload_len))
        .ok_or_else(DecodeError::projection)
}

fn varint_len(mut value: u64) -> usize {
    let mut length = 1usize;
    while value >= 128 {
        value >>= 7;
        length += 1;
    }
    length
}

fn push_varint(mut value: u64, output: &mut Vec<u8>) {
    while value >= 128 {
        output.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn check_name_rewrite_limits(
    requirements: TableModelNameRewriteRequirements,
    limits: TableModelNameRewriteLimits,
) -> Result<(), DecodeError> {
    if requirements.input_bytes > limits.input_bytes {
        return Err(DecodeError::limit(DecodeLimit::Bytes {
            observed: requirements.input_bytes,
            maximum: limits.input_bytes,
        }));
    }
    if requirements.output_bytes > limits.output_bytes {
        return Err(DecodeError::limit(DecodeLimit::Output {
            observed: requirements.output_bytes,
            maximum: limits.output_bytes,
        }));
    }
    if requirements.fields > limits.fields {
        return Err(DecodeError::limit(DecodeLimit::Fields {
            observed: requirements.fields,
            maximum: limits.fields,
        }));
    }
    if requirements.work_bytes > limits.work_bytes {
        return Err(DecodeError::limit(DecodeLimit::Work {
            observed: requirements.work_bytes,
            maximum: limits.work_bytes,
        }));
    }
    if requirements.text_bytes > limits.text_bytes {
        return Err(DecodeError::limit(DecodeLimit::Text {
            observed: requirements.text_bytes,
            maximum: limits.text_bytes,
        }));
    }
    if requirements.max_depth > limits.max_depth {
        return Err(DecodeError::limit(DecodeLimit::Nesting {
            observed: requirements.max_depth,
            maximum: limits.max_depth,
        }));
    }
    if requirements.allocations > limits.allocations {
        return Err(DecodeError::limit(DecodeLimit::Allocations {
            observed: requirements.allocations,
            maximum: limits.allocations,
        }));
    }
    if requirements.retained_bytes > limits.retained_bytes {
        return Err(DecodeError::limit(DecodeLimit::Retained {
            observed: requirements.retained_bytes,
            maximum: limits.retained_bytes,
        }));
    }
    if requirements.scratch_bytes > limits.scratch_bytes {
        return Err(DecodeError::limit(DecodeLimit::Scratch {
            observed: requirements.scratch_bytes,
            maximum: limits.scratch_bytes,
        }));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::for_source(source)
    }

    fn varint_field(number: u32, value: u64, output: &mut Vec<u8>) {
        let mut tag = u64::from(number) << 3;
        while tag >= 128 {
            output.push((tag as u8 & 0x7f) | 0x80);
            tag >>= 7;
        }
        output.push(tag as u8);
        let mut value = value;
        while value >= 128 {
            output.push((value as u8 & 0x7f) | 0x80);
            value >>= 7;
        }
        output.push(value as u8);
    }

    fn bytes_field(number: u32, value: &[u8], output: &mut Vec<u8>) {
        let mut tag = u64::from(number) << 3 | 2;
        let mut tag_bytes = [0u8; 10];
        let mut index = 0;
        while tag >= 128 {
            tag_bytes[index] = (tag as u8 & 0x7f) | 0x80;
            tag >>= 7;
            index += 1;
        }
        tag_bytes[index] = tag as u8;
        output.extend_from_slice(&tag_bytes[..=index]);
        let mut length = value.len() as u64;
        while length >= 128 {
            output.push((length as u8 & 0x7f) | 0x80);
            length >>= 7;
        }
        output.push(length as u8);
        output.extend_from_slice(value);
    }

    fn fixed64_field(number: u32, output: &mut Vec<u8>) {
        let mut tag = u64::from(number) << 3 | 1;
        let mut tag_bytes = [0u8; 10];
        let mut index = 0;
        while tag >= 128 {
            tag_bytes[index] = (tag as u8 & 0x7f) | 0x80;
            tag >>= 7;
            index += 1;
        }
        tag_bytes[index] = tag as u8;
        output.extend_from_slice(&tag_bytes[..=index]);
        output.extend_from_slice(&[0; 8]);
    }

    fn model() -> Vec<u8> {
        let mut output = Vec::new();
        bytes_field(TABLE_ID_FIELD, b"id", &mut output);
        bytes_field(TABLE_STYLE_FIELD, &[0x0a, 0x01, 0x08, 0x01], &mut output);
        bytes_field(BASE_DATA_STORE_FIELD, &[0x0a, 0x00], &mut output);
        varint_field(TABLE_ROWS_FIELD, 3, &mut output);
        varint_field(TABLE_COLUMNS_FIELD, 4, &mut output);
        bytes_field(TABLE_NAME_FIELD, b"Table", &mut output);
        fixed64_field(DEFAULT_ROW_HEIGHT_FIELD, &mut output);
        fixed64_field(DEFAULT_COLUMN_WIDTH_FIELD, &mut output);
        bytes_field(
            BODY_CELL_STYLE_FIELD,
            &[0x0a, 0x01, 0x08, 0x02],
            &mut output,
        );
        bytes_field(
            HEADER_ROW_STYLE_FIELD,
            &[0x0a, 0x01, 0x08, 0x03],
            &mut output,
        );
        bytes_field(
            HEADER_COLUMN_STYLE_FIELD,
            &[0x0a, 0x01, 0x08, 0x04],
            &mut output,
        );
        bytes_field(
            FOOTER_ROW_STYLE_FIELD,
            &[0x0a, 0x01, 0x08, 0x05],
            &mut output,
        );
        bytes_field(
            BODY_TEXT_STYLE_FIELD,
            &[0x0a, 0x01, 0x08, 0x06],
            &mut output,
        );
        bytes_field(
            HEADER_ROW_TEXT_STYLE_FIELD,
            &[0x0a, 0x01, 0x08, 0x07],
            &mut output,
        );
        bytes_field(
            HEADER_COLUMN_TEXT_STYLE_FIELD,
            &[0x0a, 0x01, 0x08, 0x08],
            &mut output,
        );
        bytes_field(
            FOOTER_ROW_TEXT_STYLE_FIELD,
            &[0x0a, 0x01, 0x08, 0x09],
            &mut output,
        );
        output
    }

    #[test]
    fn projects_required_identity_name_and_dimensions_borrowed_from_source() {
        let source = model();
        let (snapshot, report) = decode_table_model_with_report(&source, options(&source))
            .expect("strict table-model projection");
        assert_eq!(snapshot.table_id(), "id");
        assert_eq!(snapshot.table_name(), "Table");
        assert_eq!((snapshot.rows(), snapshot.columns()), (3, 4));
        assert_eq!(snapshot.number_of_rows(), 3);
        assert_eq!(snapshot.number_of_columns(), 4);
        assert_eq!(report.input_bytes(), source.len());
        assert_eq!(report.fields(), 16);
        assert_eq!(report.work_bytes(), source.len() * 3);
        assert_eq!(report.text_bytes(), 7);
        assert_eq!(report.allocations(), 0);
        assert_eq!(report.retained_bytes(), 0);
        assert_eq!(report.scratch_bytes(), 0);
        let table_name_ptr = snapshot.table_name().as_ptr() as usize;
        let source_start = source.as_ptr() as usize;
        assert!(table_name_ptr >= source_start && table_name_ptr < source_start + source.len());
    }

    #[test]
    fn canonical_unknown_scalars_fixed_values_bytes_and_groups_are_accepted() {
        let mut source = model();
        varint_field(100, 7, &mut source);
        source.extend_from_slice(&[0x99, 0x06]);
        source.extend_from_slice(&[0; 8]);
        source.extend_from_slice(&[0xa2, 0x06, 0x01, 0x7f]);
        source.extend_from_slice(&[0xab, 0x06, 0x08, 0x01, 0xac, 0x06]);
        assert!(decode_table_model(&source, options(&source)).is_ok());
    }

    #[test]
    fn missing_duplicate_wrong_wire_and_noncanonical_known_fields_fail_closed() {
        let source = model();
        let name_start = source
            .windows(2)
            .position(|window| window == [0x42, 0x05])
            .expect("table name field");
        let mut missing = source.clone();
        missing.drain(name_start..name_start + 7);
        assert_eq!(
            decode_table_model(&missing, options(&missing))
                .expect_err("missing name")
                .missing_required_field(),
            Some("TST.TableModelArchive.table_name")
        );

        let mut duplicate = source.clone();
        varint_field(TABLE_ROWS_FIELD, 9, &mut duplicate);
        assert_eq!(
            decode_table_model(&duplicate, options(&duplicate))
                .expect_err("duplicate rows")
                .duplicate_singular_field(),
            Some("TST.TableModelArchive.number_of_rows")
        );

        let mut wrong_wire = source.clone();
        wrong_wire.push(0x32);
        wrong_wire.push(1);
        wrong_wire.push(0);
        assert!(decode_table_model(&wrong_wire, options(&wrong_wire)).is_err());

        let mut noncanonical = source.clone();
        noncanonical.extend_from_slice(&[0xa0, 0x06, 0x81, 0x00]);
        assert_eq!(
            decode_table_model(&noncanonical, options(&noncanonical))
                .expect_err("noncanonical unknown scalar")
                .noncanonical_reason(),
            Some("protobuf varint value")
        );

        let mut optional_wrong_wire = source.clone();
        bytes_field(50, &[], &mut optional_wrong_wire);
        assert!(decode_table_model(&optional_wrong_wire, options(&optional_wrong_wire)).is_err());

        let mut optional_duplicate = source.clone();
        varint_field(9, 1, &mut optional_duplicate);
        varint_field(9, 2, &mut optional_duplicate);
        assert_eq!(
            decode_table_model(&optional_duplicate, options(&optional_duplicate))
                .expect_err("duplicate optional field")
                .duplicate_singular_field(),
            Some("TST.TableModelArchive.number_of_header_rows")
        );

        let mut packed_repeated = source;
        bytes_field(90, &[1], &mut packed_repeated);
        assert!(decode_table_model(&packed_repeated, options(&packed_repeated)).is_err());
    }

    #[test]
    fn limits_are_typed_and_exact_at_the_boundary() {
        let source = model();
        let exact = options(&source);
        assert!(decode_table_model(&source, exact).is_ok());
        assert_eq!(
            decode_table_model(&source, exact.with_max_input_bytes(source.len() - 1),)
                .expect_err("input bytes")
                .resource_limit(),
            Some(DecodeLimit::Bytes {
                observed: source.len(),
                maximum: source.len() - 1,
            })
        );
        assert_eq!(
            decode_table_model(&source, exact.with_max_fields(3))
                .expect_err("fields")
                .resource_limit(),
            Some(DecodeLimit::Fields {
                observed: 4,
                maximum: 3,
            })
        );
        assert_eq!(
            decode_table_model(&source, exact.with_max_work_bytes(source.len() * 3 - 1))
                .expect_err("work")
                .resource_limit(),
            Some(DecodeLimit::Work {
                observed: source.len() * 3,
                maximum: source.len() * 3 - 1,
            })
        );
        assert_eq!(
            decode_table_model(&source, exact.with_max_work_bytes(source.len() - 1))
                .expect_err("strict source-pass work")
                .resource_limit(),
            Some(DecodeLimit::Work {
                observed: source.len(),
                maximum: source.len() - 1,
            })
        );
        assert_eq!(
            decode_table_model(&source, exact.with_max_text_bytes(6))
                .expect_err("text bytes")
                .resource_limit(),
            Some(DecodeLimit::Text {
                observed: 7,
                maximum: 6,
            })
        );
        assert_eq!(
            decode_table_model(&source, exact.with_recursion_limit(0))
                .expect_err("nesting")
                .resource_limit(),
            Some(DecodeLimit::Nesting {
                observed: 0,
                maximum: MAX_RECURSION,
            })
        );
    }

    #[test]
    fn unknown_group_depth_and_malformed_framing_fail_closed() {
        let mut source = model();
        source.extend_from_slice(&[0xab, 0x06, 0xac, 0x06]);
        assert!(decode_table_model(&source, options(&source)).is_ok());

        let mut source = model();
        source.extend_from_slice(&[0xab, 0x06, 0x08, 0x01]);
        assert!(decode_table_model(&source, options(&source)).is_err());

        let mut source = model();
        source.extend_from_slice(&[0xab, 0x06, 0xab, 0x06, 0xac, 0x06, 0xac, 0x06]);
        assert_eq!(
            decode_table_model(&source, options(&source).with_recursion_limit(1))
                .expect_err("deep group")
                .resource_limit(),
            Some(DecodeLimit::Nesting {
                observed: 2,
                maximum: 1,
            })
        );

        for malformed in [
            vec![
                0x08, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x02,
            ],
            vec![0x0a, 0x80, 0x00],
            vec![0x30],
        ] {
            assert!(decode_table_model(&malformed, options(&malformed)).is_err());
        }
    }

    fn rewrite_options(source: &[u8]) -> DecodeOptions {
        let bytes = source.len().max(1);
        DecodeOptions::for_source(source)
            .with_max_input_bytes(bytes)
            .with_max_output_bytes(bytes.saturating_mul(2))
            .with_max_fields(bytes.saturating_mul(2))
            .with_max_work_bytes(bytes.saturating_mul(8))
            .with_max_text_bytes(bytes)
            .with_max_allocations(1)
            .with_max_retained_bytes(bytes.saturating_mul(2))
            .with_max_scratch_bytes(bytes.saturating_mul(3))
    }

    #[test]
    fn prepared_name_rewrite_preserves_unknowns_and_inverse_is_exact() {
        let mut source = model();
        varint_field(100, 7, &mut source);
        source.extend_from_slice(&[0xab, 0x06, 0xab, 0x06, 0xac, 0x06, 0xac, 0x06]);
        let options = rewrite_options(&source);
        let prepared = prepare_table_model_name_rewrite(&source, "Renamed", options)
            .expect("prepare table name rewrite");
        let requirements = prepared.execution_requirements();
        assert!(prepared.prepare_report().changed());
        assert!(requirements.output_bytes() > source.len());
        assert!(requirements.allocations() > 0);
        assert!(requirements.retained_bytes() > 0);
        assert!(requirements.scratch_bytes() > 0);
        let output = prepared
            .execute(requirements.exact())
            .expect("execute table name rewrite");
        assert_eq!(output.report(), prepared.prepare_report());
        assert!(
            output
                .bytes()
                .windows(2)
                .any(|window| window == [0xa0, 0x06])
        );
        assert!(
            output
                .bytes()
                .windows(8)
                .any(|window| window == [0xab, 0x06, 0xab, 0x06, 0xac, 0x06, 0xac, 0x06])
        );
        let rewritten = decode_table_model(
            output.bytes(),
            options.with_max_input_bytes(output.bytes().len()),
        )
        .expect("read rewritten table model");
        assert_eq!(rewritten.table_name(), "Renamed");
        assert_eq!((rewritten.rows(), rewritten.columns()), (3, 4));

        let inverse_options = rewrite_options(output.bytes());
        let inverse = rewrite_table_model_name(output.bytes(), "Table", inverse_options)
            .expect("inverse table name rewrite");
        assert_eq!(inverse, source);
    }

    #[test]
    fn unchanged_name_rewrite_is_an_exact_noop() {
        let source = model();
        let options = rewrite_options(&source);
        let prepared = prepare_table_model_name_rewrite(&source, "Table", options)
            .expect("prepare no-op table name rewrite");
        assert!(!prepared.prepare_report().changed());
        let output = prepared
            .execute(prepared.execution_requirements().exact())
            .expect("execute no-op table name rewrite");
        assert_eq!(output.bytes(), source.as_slice());
    }

    #[test]
    fn name_rewrite_rejects_duplicate_wrong_wire_invalid_utf8_and_stale_fingerprint() {
        let source = model();
        let options = rewrite_options(&source);

        let mut duplicate = source.clone();
        bytes_field(TABLE_NAME_FIELD, b"Other", &mut duplicate);
        let duplicate_error =
            prepare_table_model_name_rewrite(&duplicate, "Renamed", rewrite_options(&duplicate))
                .expect_err("duplicate table name");
        assert_eq!(
            duplicate_error.duplicate_singular_field(),
            Some("TST.TableModelArchive.table_name")
        );

        let mut wrong_wire = source.clone();
        varint_field(TABLE_NAME_FIELD, 1, &mut wrong_wire);
        assert!(
            prepare_table_model_name_rewrite(&wrong_wire, "Renamed", rewrite_options(&wrong_wire))
                .is_err()
        );

        let mut invalid_utf8 = source.clone();
        let name = invalid_utf8
            .windows(5)
            .position(|window| window == b"Table")
            .expect("table name payload");
        invalid_utf8[name] = 0xff;
        let utf8_error = decode_table_model(&invalid_utf8, rewrite_options(&invalid_utf8))
            .expect_err("invalid table name UTF-8");
        assert_eq!(
            utf8_error.noncanonical_reason(),
            Some("string is not valid UTF-8")
        );

        let stale = TableModelNameWrite::new("Renamed").with_fingerprint(0);
        assert!(
            prepare_table_model_name_rewrite(&source, stale, options)
                .expect_err("stale table-model fingerprint")
                .to_string()
                .contains("fingerprint")
        );
    }

    #[test]
    fn name_rewrite_execution_rejects_each_maximum_at_minus_one() {
        let mut source = model();
        source.extend_from_slice(&[0xab, 0x06, 0xab, 0x06, 0xac, 0x06, 0xac, 0x06]);
        let options = rewrite_options(&source);
        let prepared = prepare_table_model_name_rewrite(&source, "Renamed", options)
            .expect("prepare bounded table name rewrite");
        let requirements = prepared.execution_requirements();

        let assert_limit = |limits: TableModelNameRewriteLimits, expected: DecodeLimit| {
            assert_eq!(
                prepared
                    .execute(limits)
                    .expect_err("minus-one rewrite limit")
                    .resource_limit(),
                Some(expected)
            );
        };

        let maximum = requirements.input_bytes().saturating_sub(1);
        assert_limit(
            requirements.exact().with_input_bytes(maximum),
            DecodeLimit::Bytes {
                observed: requirements.input_bytes(),
                maximum,
            },
        );

        let maximum = requirements.output_bytes().saturating_sub(1);
        assert_limit(
            requirements.exact().with_output_bytes(maximum),
            DecodeLimit::Output {
                observed: requirements.output_bytes(),
                maximum,
            },
        );

        let maximum = requirements.fields().saturating_sub(1);
        assert_limit(
            requirements.exact().with_fields(maximum),
            DecodeLimit::Fields {
                observed: requirements.fields(),
                maximum,
            },
        );

        let maximum = requirements.work_bytes().saturating_sub(1);
        assert_limit(
            requirements.exact().with_work_bytes(maximum),
            DecodeLimit::Work {
                observed: requirements.work_bytes(),
                maximum,
            },
        );

        let maximum = requirements.text_bytes().saturating_sub(1);
        assert_limit(
            requirements.exact().with_text_bytes(maximum),
            DecodeLimit::Text {
                observed: requirements.text_bytes(),
                maximum,
            },
        );

        let maximum = requirements.max_depth().saturating_sub(1);
        assert_limit(
            requirements.exact().with_max_depth(maximum),
            DecodeLimit::Nesting {
                observed: requirements.max_depth(),
                maximum,
            },
        );

        let maximum = requirements.allocations().saturating_sub(1);
        assert_limit(
            requirements.exact().with_allocations(maximum),
            DecodeLimit::Allocations {
                observed: requirements.allocations(),
                maximum,
            },
        );

        let maximum = requirements.retained_bytes().saturating_sub(1);
        assert_limit(
            requirements.exact().with_retained_bytes(maximum),
            DecodeLimit::Retained {
                observed: requirements.retained_bytes(),
                maximum,
            },
        );

        let maximum = requirements.scratch_bytes().saturating_sub(1);
        assert_limit(
            requirements.exact().with_scratch_bytes(maximum),
            DecodeLimit::Scratch {
                observed: requirements.scratch_bytes(),
                maximum,
            },
        );
    }
}
