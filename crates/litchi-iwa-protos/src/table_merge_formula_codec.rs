//! Bounded Buffa authoring for a native Numbers table merge formula.
//!
//! A merge operation needs one formula payload that refers to the complete
//! rectangle.  Expanding that rectangle into one dependency per cell makes
//! encoding proportional to the rectangle's area and is especially harmful
//! for the large coordinate ranges Numbers permits.  This codec owns the
//! constant-size native representation: a `ColonTract` range node followed
//! by the `FUNCTION` node used by the Numbers host.
//!
//! The generated Buffa messages own the schema-bearing envelopes.  The two
//! small envelopes that are not represented by the existing projection
//! (`ColonTract` and its preserve-rectangle flag) are written with Buffa's
//! wire helpers.  A generated lazy-view oracle and an independent canonical
//! field walk validate every fresh payload before it is returned.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The typed request, bounded writer, and generated oracle are kept together."
)]

use std::{error::Error, fmt};

use buffa::LazyMessageView as _;

use crate::buffa_formula_generated::LitchiIwaFormulaProjection as projection;

const DEFAULT_MAX_OUTPUT_BYTES: usize = 16 * 1024;
const DEFAULT_MAX_WORK_BYTES: usize = 64 * 1024;
const DEFAULT_MAX_FIELDS: usize = 64;
const MERGE_FORMULA_FIELD_COUNT: usize = 26;
const MERGE_FORMULA_ALLOCATION_COUNT: usize = 10;

/// Finite resources for one fresh native merge formula.
///
/// `max_work_bytes` accounts for the encoded bytes of each of the ten
/// constant-size staging buffers, both generated-message size walks, the
/// generated lazy-view walk, and the independent canonical field walk. It is
/// a deterministic wire-byte ledger; it does not scale with the rectangle's
/// area. `max_fields` counts all wire fields in the nested message tree.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EncodeOptions {
    max_output_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
}

impl EncodeOptions {
    /// Construct an explicit finite policy.
    #[must_use]
    pub const fn new(max_output_bytes: usize, max_fields: usize, max_work_bytes: usize) -> Self {
        Self {
            max_output_bytes,
            max_fields,
            max_work_bytes,
        }
    }

    /// Construct the conservative default policy for one merge formula.
    #[must_use]
    pub const fn for_merge_formula() -> Self {
        Self::new(
            DEFAULT_MAX_OUTPUT_BYTES,
            DEFAULT_MAX_FIELDS,
            DEFAULT_MAX_WORK_BYTES,
        )
    }

    /// Replace the output-byte ceiling.
    #[must_use]
    pub const fn with_max_output_bytes(mut self, maximum: usize) -> Self {
        self.max_output_bytes = maximum;
        self
    }

    /// Replace the field-count ceiling.
    #[must_use]
    pub const fn with_max_fields(mut self, maximum: usize) -> Self {
        self.max_fields = maximum;
        self
    }

    /// Replace the staging-work-byte ceiling.
    #[must_use]
    pub const fn with_max_work_bytes(mut self, maximum: usize) -> Self {
        self.max_work_bytes = maximum;
        self
    }

    /// Return the output-byte ceiling.
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }

    /// Return the field-count ceiling.
    #[must_use]
    pub const fn max_fields(self) -> usize {
        self.max_fields
    }

    /// Return the staging-work-byte ceiling.
    #[must_use]
    pub const fn max_work_bytes(self) -> usize {
        self.max_work_bytes
    }
}

impl Default for EncodeOptions {
    fn default() -> Self {
        Self::for_merge_formula()
    }
}

/// Exact resource usage for a successfully encoded merge formula.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EncodeReport {
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    allocations: usize,
}

impl EncodeReport {
    /// Return the exact length of the returned formula archive.
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    /// Return the exact number of wire fields in the complete message tree.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Return the exact staged wire-byte work ledger.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Return the exact number of fixed-size staging vectors.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }
}

/// A finite resource that stopped encoding before any output was returned.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum EncodeLimit {
    /// The formula archive exceeded the output-byte ceiling.
    OutputBytes { observed: usize, maximum: usize },
    /// The nested message tree exceeded the field-count ceiling.
    Fields { observed: usize, maximum: usize },
    /// Constant-size staging exceeded the work-byte ceiling.
    WorkBytes { observed: usize, maximum: usize },
}

impl fmt::Display for EncodeLimit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::OutputBytes { observed, maximum } => write!(
                formatter,
                "merge formula output size {observed} exceeds the {maximum}-byte limit"
            ),
            Self::Fields { observed, maximum } => write!(
                formatter,
                "merge formula field count {observed} exceeds the {maximum}-field limit"
            ),
            Self::WorkBytes { observed, maximum } => write!(
                formatter,
                "merge formula work size {observed} exceeds the {maximum}-byte limit"
            ),
        }
    }
}

/// Failure while constructing a native merge formula.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum EncodeError {
    /// The supplied inclusive row or column range is inverted.
    InvalidRange,
    /// A caller-supplied finite resource was exceeded.
    Limit(EncodeLimit),
    /// A fixed-size output vector could not reserve its preflight capacity.
    Allocation { requested: usize },
    /// The generated Buffa message rejected an internal encode operation.
    Buffa(buffa::EncodeError),
    /// Checked size arithmetic overflowed or disagreed with generated Buffa.
    ArithmeticOverflow,
    /// The generated lazy-view/canonical-wire oracle rejected fresh bytes.
    GeneratedOracle,
}

impl fmt::Display for EncodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidRange => formatter.write_str("merge formula range is inverted"),
            Self::Limit(limit) => limit.fmt(formatter),
            Self::Allocation { requested } => write!(
                formatter,
                "could not reserve {requested} bytes for a merge formula stage"
            ),
            Self::Buffa(error) => write!(formatter, "Buffa merge formula encode failed: {error}"),
            Self::ArithmeticOverflow => {
                formatter.write_str("merge formula size arithmetic overflowed")
            },
            Self::GeneratedOracle => {
                formatter.write_str("generated Buffa oracle rejected merge formula bytes")
            },
        }
    }
}

impl Error for EncodeError {}

impl From<buffa::EncodeError> for EncodeError {
    fn from(error: buffa::EncodeError) -> Self {
        Self::Buffa(error)
    }
}

impl EncodeError {
    /// Return the resource failure, when encoding stopped at a finite limit.
    #[must_use]
    pub const fn limit(&self) -> Option<EncodeLimit> {
        match self {
            Self::Limit(limit) => Some(*limit),
            _ => None,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct MergeFormulaPlan {
    cfuuid_bytes: usize,
    cross_table_bytes: usize,
    sticky_bytes: usize,
    column_range_bytes: usize,
    row_range_bytes: usize,
    colon_tract_bytes: usize,
    range_node_bytes: usize,
    function_node_bytes: usize,
    ast_array_bytes: usize,
    formula_bytes: usize,
    work_bytes: usize,
}

/// Encode the canonical two-node formula used by a native Numbers table
/// merge.
///
/// Coordinates are inclusive and are kept as unsigned values all the way to
/// the generated `AbsoluteRangeArchive` messages. In particular,
/// `u32::MAX` is a valid endpoint and never causes an `end + 1` calculation.
/// The output is constant-size with respect to the rectangle's area.
///
/// The returned report is exact for the codec's resource ledger. The ten
/// allocations are the CFUUID, cross-table, sticky, two absolute-range,
/// colon-tract, range-node, function-node, AST-array, and final formula
/// buffers.
pub fn encode_merge_formula(
    table_owner_cfuuid_words: [u32; 4],
    begin_row: u32,
    end_row: u32,
    begin_column: u32,
    end_column: u32,
    options: EncodeOptions,
) -> Result<(Vec<u8>, EncodeReport), EncodeError> {
    if begin_row > end_row || begin_column > end_column {
        return Err(EncodeError::InvalidRange);
    }

    let plan = plan_merge_formula(
        table_owner_cfuuid_words,
        begin_row,
        end_row,
        begin_column,
        end_column,
    )?;
    enforce_limits(&plan, options)?;

    let owner = projection::CFUUIDArchive {
        word0: Some(table_owner_cfuuid_words[0]),
        word1: Some(table_owner_cfuuid_words[1]),
        word2: Some(table_owner_cfuuid_words[2]),
        word3: Some(table_owner_cfuuid_words[3]),
    };
    let owner_bytes = encode_message(&owner, plan.cfuuid_bytes, options)?;

    let cross_table = projection::CrossTableExtraArchive {
        table_id: owner_bytes,
    };
    let cross_table_bytes = encode_message(&cross_table, plan.cross_table_bytes, options)?;

    let sticky = projection::StickyBitsArchive {
        begin_row_absolute: true,
        begin_column_absolute: true,
        end_row_absolute: true,
        end_column_absolute: true,
    };
    let sticky_bytes = encode_message(&sticky, plan.sticky_bytes, options)?;

    let column_range = projection::AbsoluteRangeArchive {
        begin: begin_column,
        end: Some(end_column),
    };
    let column_range_bytes = encode_message(&column_range, plan.column_range_bytes, options)?;

    let row_range = projection::AbsoluteRangeArchive {
        begin: begin_row,
        end: Some(end_row),
    };
    let row_range_bytes = encode_message(&row_range, plan.row_range_bytes, options)?;

    let colon_tract_bytes = encode_colon_tract(
        &column_range_bytes,
        &row_range_bytes,
        plan.colon_tract_bytes,
    )?;

    let range_node = projection::ASTNodeArchive {
        node_type: 67,
        cross_table_extra: Some(cross_table_bytes),
        sticky_bits: Some(sticky_bytes),
        colon_tract: Some(colon_tract_bytes),
        ..Default::default()
    };
    let range_node_bytes = encode_message(&range_node, plan.range_node_bytes, options)?;

    let function_node = projection::ASTNodeArchive {
        node_type: 16,
        function_index: Some(168),
        function_num_args: Some(1),
        ..Default::default()
    };
    let function_node_bytes = encode_message(&function_node, plan.function_node_bytes, options)?;

    let ast_array_bytes = encode_ast_array(
        &range_node_bytes,
        &function_node_bytes,
        plan.ast_array_bytes,
    )?;

    let formula = projection::FormulaArchive {
        ast_node_array: ast_array_bytes,
        ..Default::default()
    };
    let output = encode_message(&formula, plan.formula_bytes, options)?;

    verify_generated_shape(
        &output,
        table_owner_cfuuid_words,
        begin_row,
        end_row,
        begin_column,
        end_column,
    )?;

    let report = EncodeReport {
        output_bytes: plan.formula_bytes,
        fields: MERGE_FORMULA_FIELD_COUNT,
        work_bytes: plan.work_bytes,
        allocations: MERGE_FORMULA_ALLOCATION_COUNT,
    };
    debug_assert_eq!(report.output_bytes, output.len());
    Ok((output, report))
}

fn plan_merge_formula(
    table_owner_cfuuid_words: [u32; 4],
    begin_row: u32,
    end_row: u32,
    begin_column: u32,
    end_column: u32,
) -> Result<MergeFormulaPlan, EncodeError> {
    let cfuuid_bytes = sum([
        varint_field_len(2, table_owner_cfuuid_words[0]),
        varint_field_len(3, table_owner_cfuuid_words[1]),
        varint_field_len(4, table_owner_cfuuid_words[2]),
        varint_field_len(5, table_owner_cfuuid_words[3]),
    ])?;
    let cross_table_bytes = bytes_field_len(1, cfuuid_bytes)?;
    let sticky_bytes = sum([
        varint_field_len(1, 1),
        varint_field_len(2, 1),
        varint_field_len(3, 1),
        varint_field_len(4, 1),
    ])?;
    let column_range_bytes = sum([
        varint_field_len(1, begin_column),
        varint_field_len(2, end_column),
    ])?;
    let row_range_bytes = sum([varint_field_len(1, begin_row), varint_field_len(2, end_row)])?;
    let colon_tract_bytes = sum([
        bytes_field_len(3, column_range_bytes)?,
        bytes_field_len(4, row_range_bytes)?,
        varint_field_len(5, 1),
    ])?;
    let range_node_bytes = sum([
        varint_field_len(1, 67),
        bytes_field_len(28, cross_table_bytes)?,
        bytes_field_len(33, sticky_bytes)?,
        bytes_field_len(40, colon_tract_bytes)?,
    ])?;
    let function_node_bytes = sum([
        varint_field_len(1, 16),
        varint_field_len(2, 168),
        varint_field_len(3, 1),
    ])?;
    let ast_array_bytes = sum([
        bytes_field_len(1, range_node_bytes)?,
        bytes_field_len(1, function_node_bytes)?,
    ])?;
    let formula_bytes = bytes_field_len(1, ast_array_bytes)?;
    let staging_bytes = sum([
        cfuuid_bytes,
        cross_table_bytes,
        sticky_bytes,
        column_range_bytes,
        row_range_bytes,
        colon_tract_bytes,
        range_node_bytes,
        function_node_bytes,
        ast_array_bytes,
        formula_bytes,
    ])?;
    // Each generated message is measured once before encoding and measured
    // again by `try_encode_bounded`; the staging sum already accounts for the
    // actual write. The canonical walk visits every stage once, while the
    // generated lazy oracle visits the generated-message stages once.
    let generated_message_bytes = sum([
        cfuuid_bytes,
        cross_table_bytes,
        sticky_bytes,
        column_range_bytes,
        row_range_bytes,
        range_node_bytes,
        function_node_bytes,
        formula_bytes,
    ])?;
    let work_bytes = sum([
        staging_bytes,
        generated_message_bytes,
        generated_message_bytes,
        staging_bytes,
        generated_message_bytes,
    ])?;

    Ok(MergeFormulaPlan {
        cfuuid_bytes,
        cross_table_bytes,
        sticky_bytes,
        column_range_bytes,
        row_range_bytes,
        colon_tract_bytes,
        range_node_bytes,
        function_node_bytes,
        ast_array_bytes,
        formula_bytes,
        work_bytes,
    })
}

fn enforce_limits(plan: &MergeFormulaPlan, options: EncodeOptions) -> Result<(), EncodeError> {
    if plan.formula_bytes > options.max_output_bytes {
        return Err(EncodeError::Limit(EncodeLimit::OutputBytes {
            observed: plan.formula_bytes,
            maximum: options.max_output_bytes,
        }));
    }
    if MERGE_FORMULA_FIELD_COUNT > options.max_fields {
        return Err(EncodeError::Limit(EncodeLimit::Fields {
            observed: MERGE_FORMULA_FIELD_COUNT,
            maximum: options.max_fields,
        }));
    }
    if plan.work_bytes > options.max_work_bytes {
        return Err(EncodeError::Limit(EncodeLimit::WorkBytes {
            observed: plan.work_bytes,
            maximum: options.max_work_bytes,
        }));
    }
    Ok(())
}

fn sum<const N: usize>(values: [usize; N]) -> Result<usize, EncodeError> {
    values.into_iter().try_fold(0usize, |total, value| {
        total
            .checked_add(value)
            .ok_or(EncodeError::ArithmeticOverflow)
    })
}

fn varint_field_len(field_number: u32, value: u32) -> usize {
    let tag = u64::from(field_number) << 3;
    buffa::encoding::varint_len(tag) + buffa::types::uint32_encoded_len(value)
}

fn bytes_field_len(field_number: u32, payload_len: usize) -> Result<usize, EncodeError> {
    let tag = (u64::from(field_number) << 3) | 2;
    let payload_len = u64::try_from(payload_len).map_err(|_| EncodeError::ArithmeticOverflow)?;
    let length = buffa::encoding::varint_len(payload_len);
    buffa::encoding::varint_len(tag)
        .checked_add(length)
        .and_then(|size| size.checked_add(payload_len as usize))
        .ok_or(EncodeError::ArithmeticOverflow)
}

fn reserve_vec(expected_len: usize) -> Result<Vec<u8>, EncodeError> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(expected_len)
        .map_err(|_| EncodeError::Allocation {
            requested: expected_len,
        })?;
    if output.capacity() < expected_len {
        return Err(EncodeError::Allocation {
            requested: expected_len,
        });
    }
    Ok(output)
}

fn encode_message<M: buffa::Message>(
    message: &M,
    expected_len: usize,
    options: EncodeOptions,
) -> Result<Vec<u8>, EncodeError> {
    let measured_len =
        usize::try_from(message.try_encoded_len()?).map_err(|_| EncodeError::ArithmeticOverflow)?;
    if measured_len != expected_len {
        return Err(EncodeError::GeneratedOracle);
    }

    let mut output = reserve_vec(expected_len)?;
    let maximum = u32::try_from(options.max_output_bytes).unwrap_or(u32::MAX);
    let written = message.try_encode_bounded(maximum, &mut output)?;
    if usize::try_from(written).ok() != Some(expected_len) || output.len() != expected_len {
        return Err(EncodeError::GeneratedOracle);
    }
    Ok(output)
}

fn encode_colon_tract(
    column_range: &[u8],
    row_range: &[u8],
    expected_len: usize,
) -> Result<Vec<u8>, EncodeError> {
    let mut output = reserve_vec(expected_len)?;
    buffa::types::put_bytes_field(3, column_range, &mut output);
    buffa::types::put_bytes_field(4, row_range, &mut output);
    buffa::types::put_bool_field(5, true, &mut output);
    if output.len() != expected_len {
        return Err(EncodeError::GeneratedOracle);
    }
    Ok(output)
}

fn encode_ast_array(
    range_node: &[u8],
    function_node: &[u8],
    expected_len: usize,
) -> Result<Vec<u8>, EncodeError> {
    let mut output = reserve_vec(expected_len)?;
    buffa::types::put_bytes_field(1, range_node, &mut output);
    buffa::types::put_bytes_field(1, function_node, &mut output);
    if output.len() != expected_len {
        return Err(EncodeError::GeneratedOracle);
    }
    Ok(output)
}

fn verify_generated_shape(
    source: &[u8],
    expected_words: [u32; 4],
    begin_row: u32,
    end_row: u32,
    begin_column: u32,
    end_column: u32,
) -> Result<(), EncodeError> {
    let options = buffa::DecodeOptions::new()
        .with_recursion_limit(8)
        .with_max_message_size(source.len())
        .with_unknown_field_limit(MERGE_FORMULA_FIELD_COUNT)
        .with_element_memory_limit(0);
    let root: projection::FormulaArchiveLazyView<'_> = options
        .decode_lazy_view(source)
        .map_err(|_| EncodeError::GeneratedOracle)?;
    if !root.has_ast_node_array()
        || root.host_column.is_some()
        || root.host_row.is_some()
        || root.host_column_is_negative.is_some()
        || root.host_row_is_negative.is_some()
        || root.translation_flags.is_some()
        || root.host_table_uid.is_some()
        || root.host_column_uid.is_some()
        || root.host_row_uid.is_some()
    {
        return Err(EncodeError::GeneratedOracle);
    }

    let mut root_fields = source;
    let ast = read_bytes(&mut root_fields, 1)?;
    if !root_fields.is_empty() || root.ast_node_array != ast {
        return Err(EncodeError::GeneratedOracle);
    }

    let mut ast_fields = ast;
    let range_node = read_bytes(&mut ast_fields, 1)?;
    let function_node = read_bytes(&mut ast_fields, 1)?;
    if !ast_fields.is_empty() {
        return Err(EncodeError::GeneratedOracle);
    }

    verify_range_node(
        range_node,
        expected_words,
        begin_row,
        end_row,
        begin_column,
        end_column,
    )?;
    verify_function_node(function_node)?;
    Ok(())
}

fn verify_range_node(
    source: &[u8],
    expected_words: [u32; 4],
    begin_row: u32,
    end_row: u32,
    begin_column: u32,
    end_column: u32,
) -> Result<(), EncodeError> {
    let mut fields = source;
    let node_type = read_u32(&mut fields, 1)?;
    let cross_table = read_bytes(&mut fields, 28)?;
    let sticky = read_bytes(&mut fields, 33)?;
    let colon = read_bytes(&mut fields, 40)?;
    if !fields.is_empty() || node_type != 67 {
        return Err(EncodeError::GeneratedOracle);
    }

    let view: projection::ASTNodeArchiveLazyView<'_> =
        projection::ASTNodeArchiveLazyView::decode_lazy(source)
            .map_err(|_| EncodeError::GeneratedOracle)?;
    if !view.has_node_type()
        || view.node_type != 67
        || view.function_index.is_some()
        || view.function_num_args.is_some()
        || view.number.is_some()
        || view.boolean.is_some()
        || view.string.is_some()
        || view.date.is_some()
        || view.duration.is_some()
        || view.token_boolean.is_some()
        || view.array_num_col.is_some()
        || view.array_num_row.is_some()
        || view.list_num_args.is_some()
        || view.thunk_array.is_some()
        || view.local_cell_reference.is_some()
        || view.cross_table_cell_reference.is_some()
        || view.unknown_function_string.is_some()
        || view.unknown_function_num_args.is_some()
        || view.whitespace.is_some()
        || view.column.is_some()
        || view.row.is_some()
        || view.uid_coordinate.is_some()
        || view.tract_list.is_some()
        || view.category_ref.is_some()
        || view.decimal_low.is_some()
        || view.decimal_high.is_some()
        || view.cross_table_extra != Some(cross_table)
        || view.sticky_bits != Some(sticky)
        || view.colon_tract != Some(colon)
    {
        return Err(EncodeError::GeneratedOracle);
    }

    verify_cross_table(cross_table, expected_words)?;
    verify_sticky(sticky)?;
    verify_colon_tract(colon, begin_row, end_row, begin_column, end_column)?;
    Ok(())
}

fn verify_function_node(source: &[u8]) -> Result<(), EncodeError> {
    let mut fields = source;
    let node_type = read_u32(&mut fields, 1)?;
    let function_index = read_u32(&mut fields, 2)?;
    let argument_count = read_u32(&mut fields, 3)?;
    if !fields.is_empty() || node_type != 16 || function_index != 168 || argument_count != 1 {
        return Err(EncodeError::GeneratedOracle);
    }

    let view: projection::ASTNodeArchiveLazyView<'_> =
        projection::ASTNodeArchiveLazyView::decode_lazy(source)
            .map_err(|_| EncodeError::GeneratedOracle)?;
    if !view.has_node_type()
        || view.node_type != 16
        || view.function_index != Some(168)
        || view.function_num_args != Some(1)
        || view.number.is_some()
        || view.boolean.is_some()
        || view.string.is_some()
        || view.date.is_some()
        || view.duration.is_some()
        || view.token_boolean.is_some()
        || view.array_num_col.is_some()
        || view.array_num_row.is_some()
        || view.list_num_args.is_some()
        || view.thunk_array.is_some()
        || view.local_cell_reference.is_some()
        || view.cross_table_cell_reference.is_some()
        || view.unknown_function_string.is_some()
        || view.unknown_function_num_args.is_some()
        || view.whitespace.is_some()
        || view.column.is_some()
        || view.row.is_some()
        || view.cross_table_extra.is_some()
        || view.uid_coordinate.is_some()
        || view.sticky_bits.is_some()
        || view.tract_list.is_some()
        || view.category_ref.is_some()
        || view.colon_tract.is_some()
        || view.decimal_low.is_some()
        || view.decimal_high.is_some()
    {
        return Err(EncodeError::GeneratedOracle);
    }
    Ok(())
}

fn verify_cross_table(source: &[u8], expected_words: [u32; 4]) -> Result<(), EncodeError> {
    let mut fields = source;
    let table_id = read_bytes(&mut fields, 1)?;
    if !fields.is_empty() {
        return Err(EncodeError::GeneratedOracle);
    }
    let view: projection::CrossTableExtraArchiveLazyView<'_> =
        projection::CrossTableExtraArchiveLazyView::decode_lazy(source)
            .map_err(|_| EncodeError::GeneratedOracle)?;
    if !view.has_table_id() || view.table_id != table_id {
        return Err(EncodeError::GeneratedOracle);
    }

    let mut uuid_fields = table_id;
    let words = [
        read_u32(&mut uuid_fields, 2)?,
        read_u32(&mut uuid_fields, 3)?,
        read_u32(&mut uuid_fields, 4)?,
        read_u32(&mut uuid_fields, 5)?,
    ];
    if !uuid_fields.is_empty() || words != expected_words {
        return Err(EncodeError::GeneratedOracle);
    }
    let uuid_view: projection::CFUUIDArchiveLazyView<'_> =
        projection::CFUUIDArchiveLazyView::decode_lazy(table_id)
            .map_err(|_| EncodeError::GeneratedOracle)?;
    if uuid_view.word0 != Some(expected_words[0])
        || uuid_view.word1 != Some(expected_words[1])
        || uuid_view.word2 != Some(expected_words[2])
        || uuid_view.word3 != Some(expected_words[3])
    {
        return Err(EncodeError::GeneratedOracle);
    }
    Ok(())
}

fn verify_sticky(source: &[u8]) -> Result<(), EncodeError> {
    let mut fields = source;
    if !read_bool(&mut fields, 1)?
        || !read_bool(&mut fields, 2)?
        || !read_bool(&mut fields, 3)?
        || !read_bool(&mut fields, 4)?
        || !fields.is_empty()
    {
        return Err(EncodeError::GeneratedOracle);
    }
    let view: projection::StickyBitsArchiveLazyView<'_> =
        projection::StickyBitsArchiveLazyView::decode_lazy(source)
            .map_err(|_| EncodeError::GeneratedOracle)?;
    if !view.has_begin_row_absolute()
        || !view.has_begin_column_absolute()
        || !view.has_end_row_absolute()
        || !view.has_end_column_absolute()
        || !view.begin_row_absolute
        || !view.begin_column_absolute
        || !view.end_row_absolute
        || !view.end_column_absolute
    {
        return Err(EncodeError::GeneratedOracle);
    }
    Ok(())
}

fn verify_colon_tract(
    source: &[u8],
    begin_row: u32,
    end_row: u32,
    begin_column: u32,
    end_column: u32,
) -> Result<(), EncodeError> {
    let mut fields = source;
    let column = read_bytes(&mut fields, 3)?;
    let row = read_bytes(&mut fields, 4)?;
    if !read_bool(&mut fields, 5)? || !fields.is_empty() {
        return Err(EncodeError::GeneratedOracle);
    }
    verify_absolute_range(column, begin_column, end_column)?;
    verify_absolute_range(row, begin_row, end_row)?;
    Ok(())
}

fn verify_absolute_range(source: &[u8], begin: u32, end: u32) -> Result<(), EncodeError> {
    let mut fields = source;
    let actual_begin = read_u32(&mut fields, 1)?;
    let actual_end = read_u32(&mut fields, 2)?;
    if !fields.is_empty() || actual_begin != begin || actual_end != end {
        return Err(EncodeError::GeneratedOracle);
    }
    let view: projection::AbsoluteRangeArchiveLazyView<'_> =
        projection::AbsoluteRangeArchiveLazyView::decode_lazy(source)
            .map_err(|_| EncodeError::GeneratedOracle)?;
    if !view.has_begin() || view.begin != begin || view.end != Some(end) {
        return Err(EncodeError::GeneratedOracle);
    }
    Ok(())
}

fn read_u32(source: &mut &[u8], field_number: u32) -> Result<u32, EncodeError> {
    let value = read_varint_field(source, field_number)?;
    u32::try_from(value).map_err(|_| EncodeError::GeneratedOracle)
}

fn read_bool(source: &mut &[u8], field_number: u32) -> Result<bool, EncodeError> {
    match read_varint_field(source, field_number)? {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(EncodeError::GeneratedOracle),
    }
}

fn read_varint_field(source: &mut &[u8], field_number: u32) -> Result<u64, EncodeError> {
    read_field_header(source, field_number, 0)?;
    read_varint(source)
}

fn read_bytes<'source>(
    source: &mut &'source [u8],
    field_number: u32,
) -> Result<&'source [u8], EncodeError> {
    read_field_header(source, field_number, 2)?;
    let length = usize::try_from(read_varint(source)?).map_err(|_| EncodeError::GeneratedOracle)?;
    if source.len() < length {
        return Err(EncodeError::GeneratedOracle);
    }
    let (payload, remaining) = source.split_at(length);
    *source = remaining;
    Ok(payload)
}

fn read_field_header(
    source: &mut &[u8],
    expected_field: u32,
    expected_wire_type: u32,
) -> Result<(), EncodeError> {
    let raw_tag = read_varint(source)?;
    let field_number = u32::try_from(raw_tag >> 3).map_err(|_| EncodeError::GeneratedOracle)?;
    let wire_type = u32::try_from(raw_tag & 0x07).map_err(|_| EncodeError::GeneratedOracle)?;
    if field_number == 0 || field_number != expected_field || wire_type != expected_wire_type {
        return Err(EncodeError::GeneratedOracle);
    }
    Ok(())
}

fn read_varint(source: &mut &[u8]) -> Result<u64, EncodeError> {
    let before = source.len();
    let mut cursor = *source;
    let value =
        buffa::encoding::decode_varint(&mut cursor).map_err(|_| EncodeError::GeneratedOracle)?;
    let consumed = before
        .checked_sub(cursor.len())
        .ok_or(EncodeError::GeneratedOracle)?;
    if consumed != buffa::encoding::varint_len(value) {
        return Err(EncodeError::GeneratedOracle);
    }
    *source = cursor;
    Ok(value)
}

#[cfg(test)]
mod tests {
    use super::{EncodeError, EncodeLimit, EncodeOptions, encode_merge_formula};

    const OWNER: [u32; 4] = [1, 2, 3, 4];

    fn encode() -> (Vec<u8>, super::EncodeReport) {
        encode_merge_formula(OWNER, 2, 9, 3, 8, EncodeOptions::default()).unwrap()
    }

    #[test]
    fn emits_constant_size_native_formula() {
        let (output, report) = encode();
        // Golden bytes from the native Numbers TSCE/TSP writer for the same
        // owner and inclusive rectangle. This catches field ordering and the
        // preserve-rectangle flag independently of the generated oracle.
        assert_eq!(
            output.as_slice(),
            [
                0x0a, 0x36, 0x0a, 0x2b, 0x08, 0x43, 0xe2, 0x01, 0x0a, 0x0a, 0x08, 0x10, 0x01, 0x18,
                0x02, 0x20, 0x03, 0x28, 0x04, 0x8a, 0x02, 0x08, 0x08, 0x01, 0x10, 0x01, 0x18, 0x01,
                0x20, 0x01, 0xc2, 0x02, 0x0e, 0x1a, 0x04, 0x08, 0x03, 0x10, 0x08, 0x22, 0x04, 0x08,
                0x02, 0x10, 0x09, 0x28, 0x01, 0x0a, 0x07, 0x08, 0x10, 0x10, 0xa8, 0x01, 0x18, 0x01,
            ]
        );
        assert_eq!(report.output_bytes(), output.len());
        assert_eq!(report.fields(), 26);
        assert_eq!(report.allocations(), 10);
        assert_eq!(report.work_bytes(), 836);

        let (huge, huge_report) = encode_merge_formula(
            OWNER,
            u32::MAX - 3,
            u32::MAX,
            0,
            u32::MAX,
            EncodeOptions::default(),
        )
        .unwrap();
        assert!(huge.len() < 128);
        assert_eq!(huge_report.fields(), report.fields());
        assert!(huge_report.work_bytes() > report.work_bytes());
    }

    #[test]
    fn rejects_inverted_ranges_before_allocation() {
        let result = encode_merge_formula(OWNER, 4, 3, 0, 0, EncodeOptions::default());
        assert!(matches!(result, Err(EncodeError::InvalidRange)));
    }

    #[test]
    fn reports_each_finite_limit_exactly() {
        let (_, report) = encode();
        let output = encode_merge_formula(
            OWNER,
            2,
            9,
            3,
            8,
            EncodeOptions::default().with_max_output_bytes(report.output_bytes() - 1),
        );
        assert!(matches!(
            output,
            Err(EncodeError::Limit(EncodeLimit::OutputBytes { .. }))
        ));

        let fields = encode_merge_formula(
            OWNER,
            2,
            9,
            3,
            8,
            EncodeOptions::default().with_max_fields(25),
        );
        assert!(matches!(
            fields,
            Err(EncodeError::Limit(EncodeLimit::Fields { .. }))
        ));

        let work = encode_merge_formula(
            OWNER,
            2,
            9,
            3,
            8,
            EncodeOptions::default().with_max_work_bytes(report.work_bytes() - 1),
        );
        assert!(matches!(
            work,
            Err(EncodeError::Limit(EncodeLimit::WorkBytes { .. }))
        ));
    }
}
