//! Strict, generated-free table-appearance wire ownership.
//!
//! The Numbers table appearance route touches three small protobuf envelopes:
//! `TST.TableModelArchive.table_style`, `TST.TableStyleArchive` (including
//! its `TSS.StyleArchive` parent edge and `TST.TableStylePropertiesArchive`),
//! and `TSS.StylesheetArchive`.  This module deliberately projects only the
//! scalar facts needed by a package owner.  Every rewrite copies untouched
//! source fields verbatim, including unknown groups and overlong unknown
//! scalar values.  Known selected fields use canonical keys, lengths, and
//! scalar encodings.
//!
//! Object/archive lookup, UUIDs, metadata, and package publication stay above
//! this codec.  No generated schema type crosses this boundary.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The strict wire pass intentionally precedes the small public projections."
)]

use core::fmt;

const MODEL_STYLE_FIELD: u32 = 3;
const MODEL_STYLE_PRESET_FIELD: u32 = 48;
const MODEL_STYLE_REFERENCE_FIELD: u32 = 1;

const STYLE_SUPER_FIELD: u32 = 1;
const STYLE_NAME_FIELD: u32 = 1;
const STYLE_IDENTIFIER_FIELD: u32 = 2;
const STYLE_PARENT_FIELD: u32 = 3;
const STYLE_VARIATION_FIELD: u32 = 4;
const STYLE_STYLESHEET_FIELD: u32 = 5;
const STYLE_OVERRIDE_COUNT_FIELD: u32 = 10;
const STYLE_PROPERTIES_FIELD: u32 = 11;

const PROPERTIES_BANDED_ROWS_FIELD: u32 = 1;
const PROPERTIES_AUTO_RESIZE_FIELD: u32 = 22;
const PROPERTIES_BODY_HORIZONTAL_FIELD: u32 = 33;
const PROPERTIES_BODY_VERTICAL_FIELD: u32 = 34;
const PROPERTIES_LEGACY_HEADER_ROW_FIELD: u32 = 35;
const PROPERTIES_LEGACY_HEADER_COLUMN_FIELD: u32 = 36;
const PROPERTIES_LEGACY_FOOTER_ROW_FIELD: u32 = 37;
const PROPERTIES_HEADER_COLUMN_FIELD: u32 = 42;
const PROPERTIES_HEADER_ROW_FIELD: u32 = 43;
const PROPERTIES_FOOTER_ROW_FIELD: u32 = 44;

const SHEET_STYLES_FIELD: u32 = 1;
const SHEET_IDENTIFIED_STYLES_FIELD: u32 = 2;
const SHEET_PARENT_FIELD: u32 = 3;
const SHEET_CHILDREN_FIELD: u32 = 5;
const SHEET_VERSIONED_FIRST_FIELD: u32 = 7;
const SHEET_VERSIONED_LAST_FIELD: u32 = 22;
const IDENTIFIED_IDENTIFIER_FIELD: u32 = 1;
const IDENTIFIED_STYLE_FIELD: u32 = 2;
const CHILDREN_PARENT_FIELD: u32 = 1;
const CHILDREN_STYLE_FIELD: u32 = 2;
const VERSIONED_STYLES_FIELD: u32 = 1;
const VERSIONED_IDENTIFIED_STYLES_FIELD: u32 = 2;
const VERSIONED_CHILDREN_FIELD: u32 = 3;

const MAX_FIELD_NUMBER: u32 = 0x1fff_ffff;
const MAX_DEFAULT_RECURSION: u32 = 64;
// Registry facts are kept in fixed stack storage so prepare never allocates.
// Native Numbers stylesheets routinely contain more than 256 registered
// styles, so the codec admits a bounded producer-sized registry while callers
// may still choose any lower `max_styles` ceiling.
const MAX_REGISTRY_FACTS: usize = 1_024;

/// Finite limits for one source decode or prepared rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_input_bytes: usize,
    max_output_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
    max_styles: usize,
    max_allocations: usize,
}

impl DecodeOptions {
    /// Construct a finite byte/field/work/nesting/style policy.
    #[must_use]
    pub const fn new(
        max_input_bytes: usize,
        max_output_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        recursion_limit: u32,
        max_styles: usize,
    ) -> Self {
        Self {
            max_input_bytes,
            max_output_bytes,
            max_fields,
            max_work_bytes,
            recursion_limit,
            max_styles,
            max_allocations: 64,
        }
    }

    /// Build a conservative finite policy for one caller-owned payload.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        Self::new(
            bytes,
            bytes.saturating_mul(2).max(1),
            bytes.saturating_mul(8).max(1),
            bytes.saturating_mul(16).max(1),
            16,
            bytes.max(1),
        )
    }

    #[must_use]
    pub const fn with_max_output_bytes(mut self, value: usize) -> Self {
        self.max_output_bytes = value;
        self
    }

    #[must_use]
    pub const fn with_max_input_bytes(mut self, value: usize) -> Self {
        self.max_input_bytes = value;
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
    pub const fn with_max_styles(mut self, value: usize) -> Self {
        self.max_styles = value;
        self
    }

    #[must_use]
    pub const fn with_max_allocations(mut self, value: usize) -> Self {
        self.max_allocations = value;
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
    pub const fn max_styles(self) -> usize {
        self.max_styles
    }

    #[must_use]
    pub const fn max_allocations(self) -> usize {
        self.max_allocations
    }
}

/// Alias retained for package callers that use rewrite-specific terminology.
pub type RewriteOptions = DecodeOptions;

/// Typed finite failure from strict appearance parsing or rewriting.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    /// Source bytes exceeded the input ceiling.
    InputBytes { observed: usize, maximum: usize },
    /// Candidate bytes exceeded the output ceiling.
    OutputBytes { observed: usize, maximum: usize },
    /// Strict field visits exceeded the aggregate ceiling.
    Fields { observed: usize, maximum: usize },
    /// Traversal/assembly work exceeded the aggregate ceiling.
    WorkBytes { observed: usize, maximum: usize },
    /// Protobuf/group nesting exceeded the configured ceiling.
    Nesting { observed: u32, maximum: u32 },
    /// Style registry entries exceeded the configured ceiling.
    Styles { observed: usize, maximum: usize },
    /// Output allocation count exceeded the configured ceiling.
    Allocations { observed: usize, maximum: usize },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ErrorKind {
    Invalid(&'static str),
    Limit(DecodeLimit),
    Allocation { requested: usize },
}

/// Strict appearance codec failure. Diagnostics contain schema labels only;
/// source bytes and authored identifiers are never formatted.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeError {
    kind: ErrorKind,
}

impl DecodeError {
    const fn invalid(reason: &'static str) -> Self {
        Self {
            kind: ErrorKind::Invalid(reason),
        }
    }

    const fn limit(limit: DecodeLimit) -> Self {
        Self {
            kind: ErrorKind::Limit(limit),
        }
    }

    const fn allocation(requested: usize) -> Self {
        Self {
            kind: ErrorKind::Allocation { requested },
        }
    }

    /// Return the typed finite resource failure, if applicable.
    #[must_use]
    pub const fn resource_limit(self) -> Option<DecodeLimit> {
        match self.kind {
            ErrorKind::Limit(limit) => Some(limit),
            _ => None,
        }
    }

    /// Return an output allocation request refused by the codec.
    #[must_use]
    pub const fn allocation_requested(self) -> Option<usize> {
        match self.kind {
            ErrorKind::Allocation { requested } => Some(requested),
            _ => None,
        }
    }

    /// Whether the failure is a malformed/semantically invalid source.
    #[must_use]
    pub const fn is_invalid(self) -> bool {
        matches!(self.kind, ErrorKind::Invalid(_))
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self.kind {
            ErrorKind::Invalid(reason) => {
                write!(formatter, "invalid table appearance wire: {reason}")
            },
            ErrorKind::Limit(DecodeLimit::InputBytes { observed, maximum }) => write!(
                formatter,
                "table appearance input is {observed} bytes; maximum is {maximum}"
            ),
            ErrorKind::Limit(DecodeLimit::OutputBytes { observed, maximum }) => write!(
                formatter,
                "table appearance output is {observed} bytes; maximum is {maximum}"
            ),
            ErrorKind::Limit(DecodeLimit::Fields { observed, maximum }) => write!(
                formatter,
                "table appearance visited {observed} fields; maximum is {maximum}"
            ),
            ErrorKind::Limit(DecodeLimit::WorkBytes { observed, maximum }) => write!(
                formatter,
                "table appearance requires {observed} work bytes; maximum is {maximum}"
            ),
            ErrorKind::Limit(DecodeLimit::Nesting { observed, maximum }) => write!(
                formatter,
                "table appearance nesting is {observed}; maximum is {maximum}"
            ),
            ErrorKind::Limit(DecodeLimit::Styles { observed, maximum }) => write!(
                formatter,
                "table appearance registry has {observed} styles; maximum is {maximum}"
            ),
            ErrorKind::Limit(DecodeLimit::Allocations { observed, maximum }) => write!(
                formatter,
                "table appearance requested {observed} allocations; maximum is {maximum}"
            ),
            ErrorKind::Allocation { requested } => {
                write!(
                    formatter,
                    "table appearance output allocation of {requested} bytes failed"
                )
            },
        }
    }
}

impl std::error::Error for DecodeError {}

/// Aggregate strict consumption for a decode or rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct DecodeReport {
    input_bytes: usize,
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    allocations: usize,
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
    pub const fn allocations(self) -> usize {
        self.allocations
    }
}

/// Limits used by a prepared plan's output phase.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionLimits {
    pub max_input_bytes: usize,
    pub max_output_bytes: usize,
    pub max_fields: usize,
    pub max_work_bytes: usize,
    pub max_depth: u32,
    pub max_allocations: usize,
}

impl RewriteExecutionLimits {
    #[must_use]
    pub const fn new(
        max_input_bytes: usize,
        max_output_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        max_depth: u32,
        max_allocations: usize,
    ) -> Self {
        Self {
            max_input_bytes,
            max_output_bytes,
            max_fields,
            max_work_bytes,
            max_depth,
            max_allocations,
        }
    }

    #[must_use]
    pub const fn exact(report: DecodeReport) -> Self {
        Self::new(
            report.input_bytes,
            report.output_bytes,
            report.fields,
            report.work_bytes,
            report.max_depth,
            report.allocations,
        )
    }
}

/// Exact preflight requirements. No candidate output allocation occurs while
/// computing this value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionRequirements {
    input_bytes: usize,
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    allocations: usize,
}

impl RewriteExecutionRequirements {
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
    pub const fn allocations(self) -> usize {
        self.allocations
    }
    #[must_use]
    pub const fn exact_limits(self) -> RewriteExecutionLimits {
        RewriteExecutionLimits::new(
            self.input_bytes,
            self.output_bytes,
            self.fields,
            self.work_bytes,
            self.max_depth,
            self.allocations,
        )
    }
}

/// Requested direct semantic appearance values. `None` means inherit from a
/// parent style; `Some` means emit a canonical explicit bool in a variation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct AppearanceOverrides {
    pub row_banding: Option<bool>,
    pub row_sizing: Option<bool>,
    pub body_horizontal: Option<bool>,
    pub body_vertical: Option<bool>,
    pub header_columns_horizontal: Option<bool>,
    pub header_rows_vertical: Option<bool>,
    pub footer_rows_vertical: Option<bool>,
}

/// Effective appearance after style-parent inheritance.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct AppearanceSnapshot {
    pub row_banding: bool,
    pub row_sizing: bool,
    pub body_horizontal: bool,
    pub body_vertical: bool,
    pub header_columns_horizontal: bool,
    pub header_rows_vertical: bool,
    pub footer_rows_vertical: bool,
}

impl Default for AppearanceSnapshot {
    fn default() -> Self {
        Self {
            row_banding: false,
            row_sizing: false,
            body_horizontal: true,
            body_vertical: true,
            header_columns_horizontal: true,
            header_rows_vertical: true,
            footer_rows_vertical: true,
        }
    }
}

/// Borrowed TableModel style-edge projection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableModelSnapshot<'source> {
    source: &'source [u8],
    style_identifier: u64,
    style_preset_identifier: Option<u64>,
}

impl<'source> TableModelSnapshot<'source> {
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.source
    }
    #[must_use]
    pub const fn style_identifier(self) -> u64 {
        self.style_identifier
    }
    #[must_use]
    pub const fn style_preset_identifier(self) -> Option<u64> {
        self.style_preset_identifier
    }
}

/// Borrowed TableStyle parent/stylesheet and direct-override projection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableStyleSnapshot<'source> {
    source: &'source [u8],
    style_identifier: Option<&'source str>,
    parent_identifier: Option<u64>,
    stylesheet_identifier: Option<u64>,
    is_variation: bool,
    overrides: AppearanceOverrides,
}

impl<'source> TableStyleSnapshot<'source> {
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.source
    }
    #[must_use]
    pub const fn style_identifier(self) -> Option<&'source str> {
        self.style_identifier
    }
    #[must_use]
    pub const fn parent_identifier(self) -> Option<u64> {
        self.parent_identifier
    }
    #[must_use]
    pub const fn stylesheet_identifier(self) -> Option<u64> {
        self.stylesheet_identifier
    }
    #[must_use]
    pub const fn is_variation(self) -> bool {
        self.is_variation
    }
    #[must_use]
    pub const fn overrides(self) -> AppearanceOverrides {
        self.overrides
    }
}

/// A style node used by the allocation-free inheritance resolver.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableStyleNode<'source> {
    identifier: u64,
    style: TableStyleSnapshot<'source>,
}

impl<'source> TableStyleNode<'source> {
    #[must_use]
    pub const fn new(identifier: u64, style: TableStyleSnapshot<'source>) -> Self {
        Self { identifier, style }
    }
    #[must_use]
    pub const fn identifier(self) -> u64 {
        self.identifier
    }
    #[must_use]
    pub const fn style(self) -> TableStyleSnapshot<'source> {
        self.style
    }
}

/// Borrowed stylesheet registry facts.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct StylesheetSnapshot<'source> {
    source: &'source [u8],
    style_count: usize,
    parent_identifier: Option<u64>,
}

impl<'source> StylesheetSnapshot<'source> {
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.source
    }
    #[must_use]
    pub const fn style_count(self) -> usize {
        self.style_count
    }
    #[must_use]
    pub const fn parent_identifier(self) -> Option<u64> {
        self.parent_identifier
    }
}

/// Canonical style variation request. The package owner supplies fresh native
/// identifiers and owns UUID/metadata registration around the returned bytes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableStyleVariationWrite {
    pub parent_identifier: u64,
    pub stylesheet_identifier: u64,
    pub overrides: AppearanceOverrides,
}

/// Owned canonical TableStyleArchive payload.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableStyleVariationPayload {
    bytes: Vec<u8>,
    report: DecodeReport,
}

impl TableStyleVariationPayload {
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
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

/// Request to append one style and its identified/parent-child registry edges.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct StylesheetStyleAppend {
    pub style_identifier: u64,
    pub parent_identifier: Option<u64>,
}

/// Exact report for any prepared rewrite.
pub type RewriteReport = DecodeReport;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum RewriteKind {
    ModelStyle {
        old: u64,
        new: u64,
    },
    StylesheetAppend {
        append: StylesheetStyleAppend,
        parent_entry_exists: bool,
    },
}

/// Prepared model/style-edge or stylesheet append plan. It borrows source and
/// does not allocate the candidate output until `execute` is called.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PreparedRewrite<'source> {
    source: &'source [u8],
    kind: RewriteKind,
    candidate_max_styles: usize,
    prepare_report: DecodeReport,
    requirements: RewriteExecutionRequirements,
}

impl<'source> PreparedRewrite<'source> {
    #[must_use]
    pub const fn prepare_report(&self) -> DecodeReport {
        self.prepare_report
    }
    #[must_use]
    pub const fn execution_requirements(&self) -> RewriteExecutionRequirements {
        self.requirements
    }

    /// Execute after all byte/field/work/allocation ceilings are checked.
    pub fn execute(
        self,
        limits: RewriteExecutionLimits,
    ) -> Result<(Vec<u8>, RewriteReport), DecodeError> {
        validate_requirements(self.requirements, limits)?;
        let output_len = self.requirements.output_bytes;
        let mut output = Vec::new();
        output
            .try_reserve_exact(output_len)
            .map_err(|_| DecodeError::allocation(output_len))?;
        output.resize(output_len, 0);
        let assembly_work = self
            .source
            .len()
            .checked_add(output_len)
            .ok_or_else(|| DecodeError::invalid("work overflow"))?;
        match self.kind {
            RewriteKind::ModelStyle { old, new } => {
                rewrite_model_into(self.source, old, new, &mut output)?;
            },
            RewriteKind::StylesheetAppend {
                append,
                parent_entry_exists,
            } => {
                append_stylesheet_into(self.source, append, parent_entry_exists, &mut output)?;
            },
        }
        let (candidate_fields, candidate_work, candidate_depth) = match self.kind {
            RewriteKind::ModelStyle { .. } => {
                let candidate = scan_model(
                    &output,
                    limits.max_fields,
                    limits.max_work_bytes,
                    limits.max_depth,
                )?;
                (candidate.fields, candidate.work_bytes, candidate.max_depth)
            },
            RewriteKind::StylesheetAppend { .. } => {
                let candidate = scan_stylesheet(
                    &output,
                    limits.max_fields,
                    limits.max_work_bytes,
                    limits.max_depth,
                    self.candidate_max_styles,
                    None,
                )?
                .0;
                (candidate.fields, candidate.work_bytes, candidate.max_depth)
            },
        };
        let observed_fields = self.prepare_report.fields.saturating_add(candidate_fields);
        let observed_work = self
            .prepare_report
            .work_bytes
            .saturating_add(assembly_work)
            .saturating_add(candidate_work);
        if observed_fields > self.requirements.fields {
            return Err(DecodeError::limit(DecodeLimit::Fields {
                observed: observed_fields,
                maximum: self.requirements.fields,
            }));
        }
        if observed_work > self.requirements.work_bytes {
            return Err(DecodeError::limit(DecodeLimit::WorkBytes {
                observed: observed_work,
                maximum: limits.max_work_bytes,
            }));
        }
        let report = DecodeReport {
            input_bytes: self.requirements.input_bytes,
            output_bytes: output_len,
            fields: self.requirements.fields,
            work_bytes: self.requirements.work_bytes,
            max_depth: self.requirements.max_depth.max(candidate_depth),
            allocations: self.requirements.allocations,
        };
        Ok((output, report))
    }
}

/// Prepared TableModel style-edge rewrite alias for package-facing names.
pub type PreparedTableModelStyleRewrite<'source> = PreparedRewrite<'source>;
/// Prepared stylesheet registry append alias for package-facing names.
pub type PreparedStylesheetAppend<'source> = PreparedRewrite<'source>;

/// Decode the selected TableModel style edges.
pub fn decode_table_model(
    source: &[u8],
    options: DecodeOptions,
) -> Result<TableModelSnapshot<'_>, DecodeError> {
    Ok(decode_table_model_with_report(source, options)?.0)
}

/// Decode the selected TableModel style edges with exact source accounting.
pub fn decode_table_model_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(TableModelSnapshot<'_>, DecodeReport), DecodeError> {
    ensure_input(source, options)?;
    let scan = scan_model(
        source,
        options.max_fields,
        options.max_work_bytes,
        options.recursion_limit,
    )?;
    Ok((
        TableModelSnapshot {
            source,
            style_identifier: scan.style_identifier,
            style_preset_identifier: scan.style_preset_identifier,
        },
        scan.report(source.len(), 0),
    ))
}

/// Prepare a source-preserving replacement of `TableModelArchive.table_style`.
pub fn prepare_table_model_style_rewrite<'source>(
    source: &'source [u8],
    old_style_identifier: u64,
    new_style_identifier: u64,
    options: DecodeOptions,
) -> Result<PreparedTableModelStyleRewrite<'source>, DecodeError> {
    ensure_input(source, options)?;
    let scan = scan_model(
        source,
        options.max_fields,
        options.max_work_bytes,
        options.recursion_limit,
    )?;
    if scan.style_identifier != old_style_identifier {
        return Err(DecodeError::invalid(
            "table model style edge does not match expected source",
        ));
    }
    let old_len = encoded_varint_len(old_style_identifier);
    let new_len = encoded_varint_len(new_style_identifier);
    let output_bytes = if new_len >= old_len {
        source
            .len()
            .checked_add(new_len - old_len)
            .ok_or_else(|| DecodeError::invalid("output overflow"))?
    } else {
        source
            .len()
            .checked_sub(old_len - new_len)
            .ok_or_else(|| DecodeError::invalid("output underflow"))?
    };
    let candidate_work = output_bytes
        .checked_add(scan.nested_reference_bytes(new_len))
        .ok_or_else(|| DecodeError::invalid("work overflow"))?;
    let requirements = RewriteExecutionRequirements {
        input_bytes: source.len(),
        output_bytes,
        fields: scan.fields.saturating_mul(2),
        work_bytes: scan
            .work_bytes
            .saturating_mul(2)
            .saturating_add(source.len())
            .saturating_add(candidate_work)
            .saturating_add(output_bytes),
        max_depth: scan.max_depth,
        // The raw-preserving encoder uses a bounded set of nested/field
        // scratch buffers before the single candidate Vec is returned.  This
        // is deliberately conservative; callers can lower the ceiling and
        // receive a typed allocation-limit failure before execute allocates.
        allocations: 16,
    };
    validate_requirements_against_options(requirements, options)?;
    Ok(PreparedRewrite {
        source,
        kind: RewriteKind::ModelStyle {
            old: old_style_identifier,
            new: new_style_identifier,
        },
        candidate_max_styles: options.max_styles,
        prepare_report: scan.report(source.len(), 0),
        requirements,
    })
}

/// One-shot TableModel style-edge replacement.
pub fn rewrite_table_model_style(
    source: &[u8],
    old_style_identifier: u64,
    new_style_identifier: u64,
    options: DecodeOptions,
) -> Result<(Vec<u8>, RewriteReport), DecodeError> {
    let prepared = prepare_table_model_style_rewrite(
        source,
        old_style_identifier,
        new_style_identifier,
        options,
    )?;
    let limits = prepared.execution_requirements().exact_limits();
    prepared.execute(limits)
}

/// Decode one TableStyleArchive and its direct appearance overrides.
pub fn decode_table_style(
    source: &[u8],
    options: DecodeOptions,
) -> Result<TableStyleSnapshot<'_>, DecodeError> {
    Ok(decode_table_style_with_report(source, options)?.0)
}

/// Decode one TableStyleArchive and return exact strict consumption.
pub fn decode_table_style_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(TableStyleSnapshot<'_>, DecodeReport), DecodeError> {
    ensure_input(source, options)?;
    let mut budget = Budget::new(options);
    let snapshot = parse_table_style(source, 1, &mut budget)?;
    Ok((
        TableStyleSnapshot {
            source,
            style_identifier: snapshot.style_identifier,
            parent_identifier: snapshot.parent_identifier,
            stylesheet_identifier: snapshot.stylesheet_identifier,
            is_variation: snapshot.is_variation,
            overrides: snapshot.overrides,
        },
        budget.report(source.len(), 0),
    ))
}

/// Resolve direct TableStyle overrides through a bounded parent chain.
pub fn resolve_table_style_appearance(
    nodes: &[TableStyleNode<'_>],
    first_identifier: u64,
    options: DecodeOptions,
) -> Result<AppearanceSnapshot, DecodeError> {
    if nodes.len() > options.max_styles {
        return Err(DecodeError::limit(DecodeLimit::Styles {
            observed: nodes.len(),
            maximum: options.max_styles,
        }));
    }
    let mut budget = Budget::new(options);
    let duplicate_scan_work = nodes
        .len()
        .checked_mul(nodes.len())
        .ok_or_else(|| DecodeError::invalid("style node scan overflow"))?;
    budget.charge(duplicate_scan_work, 1)?;
    for (index, node) in nodes.iter().enumerate() {
        if nodes[..index]
            .iter()
            .any(|previous| previous.identifier == node.identifier)
        {
            return Err(DecodeError::invalid(
                "duplicate table style node identifier",
            ));
        }
    }
    let mut seen = [0u64; 64];
    let mut seen_len = 0usize;
    let mut current = Some(first_identifier);
    let mut result = AppearanceOverrides::default();
    let mut depth = 1;
    while depth <= options.recursion_limit {
        let Some(identifier) = current else { break };
        if seen[..seen_len].contains(&identifier) {
            return Err(DecodeError::invalid("table style inheritance cycle"));
        }
        if seen_len == seen.len() {
            return Err(DecodeError::limit(DecodeLimit::Nesting {
                observed: depth,
                maximum: options.recursion_limit,
            }));
        }
        seen[seen_len] = identifier;
        seen_len += 1;
        let node = nodes
            .iter()
            .find(|node| node.identifier == identifier)
            .ok_or_else(|| DecodeError::invalid("table style parent is missing"))?;
        merge_overrides(&mut result, node.style.overrides);
        budget.charge_fields(1, depth)?;
        current = node.style.parent_identifier;
        depth = depth.saturating_add(1);
    }
    if current.is_some() {
        return Err(DecodeError::limit(DecodeLimit::Nesting {
            observed: options.recursion_limit.saturating_add(1),
            maximum: options.recursion_limit,
        }));
    }
    Ok(AppearanceSnapshot {
        row_banding: result.row_banding.unwrap_or(false),
        row_sizing: result.row_sizing.unwrap_or(false),
        body_horizontal: result.body_horizontal.unwrap_or(true),
        body_vertical: result.body_vertical.unwrap_or(true),
        header_columns_horizontal: result.header_columns_horizontal.unwrap_or(true),
        header_rows_vertical: result.header_rows_vertical.unwrap_or(true),
        footer_rows_vertical: result.footer_rows_vertical.unwrap_or(true),
    })
}

/// Build a canonical variation payload without reconstructing any existing
/// source. Unknown fields are not applicable to fresh data; package owners
/// retain all old style bytes separately.
pub fn canonical_table_style_variation(
    write: TableStyleVariationWrite,
    options: DecodeOptions,
) -> Result<TableStyleVariationPayload, DecodeError> {
    if !complete_overrides(write.overrides) {
        return Err(DecodeError::invalid(
            "canonical variation requires all appearance overrides",
        ));
    }
    let properties_len = canonical_properties_len(write.overrides);
    let reference_len = |identifier: u64| varint_field_len(MODEL_STYLE_REFERENCE_FIELD, identifier);
    let parent_reference_len = reference_len(write.parent_identifier);
    let stylesheet_reference_len = reference_len(write.stylesheet_identifier);
    let parent_field_len = length_field_len(STYLE_PARENT_FIELD, parent_reference_len);
    let stylesheet_field_len = length_field_len(STYLE_STYLESHEET_FIELD, stylesheet_reference_len);
    let super_len = parent_field_len
        .checked_add(varint_field_len(STYLE_VARIATION_FIELD, 1))
        .and_then(|length| length.checked_add(stylesheet_field_len))
        .ok_or_else(|| DecodeError::invalid("variation size overflow"))?;
    let output_len = length_field_len(STYLE_SUPER_FIELD, super_len)
        .checked_add(varint_field_len(STYLE_OVERRIDE_COUNT_FIELD, 7))
        .and_then(|length| {
            length.checked_add(length_field_len(STYLE_PROPERTIES_FIELD, properties_len))
        })
        .ok_or_else(|| DecodeError::invalid("variation size overflow"))?;
    // The strict projection visits three root fields, five fields in the
    // StyleArchive super (including the two nested Reference fields), and
    // seven property scalars.  Work is the root bytes plus the nested super,
    // property, and reference field bytes visited again by those scans.
    let fields = 15usize;
    let work = output_len
        .checked_add(super_len)
        .and_then(|work| work.checked_add(properties_len))
        .and_then(|work| work.checked_add(parent_reference_len))
        .and_then(|work| work.checked_add(stylesheet_reference_len))
        .ok_or_else(|| DecodeError::invalid("variation work overflow"))?;
    if output_len > options.max_output_bytes {
        return Err(DecodeError::limit(DecodeLimit::OutputBytes {
            observed: output_len,
            maximum: options.max_output_bytes,
        }));
    }
    if fields > options.max_fields {
        return Err(DecodeError::limit(DecodeLimit::Fields {
            observed: fields,
            maximum: options.max_fields,
        }));
    }
    if work > options.max_work_bytes {
        return Err(DecodeError::limit(DecodeLimit::WorkBytes {
            observed: work,
            maximum: options.max_work_bytes,
        }));
    }
    if 3 > options.recursion_limit {
        return Err(DecodeError::limit(DecodeLimit::Nesting {
            observed: 3,
            maximum: options.recursion_limit,
        }));
    }
    if 16 > options.max_allocations {
        return Err(DecodeError::limit(DecodeLimit::Allocations {
            observed: 16,
            maximum: options.max_allocations,
        }));
    }
    let properties = canonical_properties_fallible(write.overrides)?;
    let style_super = {
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(super_len)
            .map_err(|_| DecodeError::allocation(super_len))?;
        let parent_reference = reference_payload_fallible(write.parent_identifier)?;
        bytes.extend_from_slice(&length_field_fallible(
            STYLE_PARENT_FIELD,
            &parent_reference,
        )?);
        bytes.extend_from_slice(&bool_field_fallible(STYLE_VARIATION_FIELD, true)?);
        let stylesheet_reference = reference_payload_fallible(write.stylesheet_identifier)?;
        bytes.extend_from_slice(&length_field_fallible(
            STYLE_STYLESHEET_FIELD,
            &stylesheet_reference,
        )?);
        bytes
    };
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(output_len)
        .map_err(|_| DecodeError::allocation(output_len))?;
    bytes.extend_from_slice(&length_field_fallible(STYLE_SUPER_FIELD, &style_super)?);
    bytes.extend_from_slice(&varint_field_fallible(STYLE_OVERRIDE_COUNT_FIELD, 7)?);
    bytes.extend_from_slice(&length_field_fallible(STYLE_PROPERTIES_FIELD, &properties)?);
    debug_assert_eq!(bytes.len(), output_len);
    let report = DecodeReport {
        output_bytes: output_len,
        fields,
        work_bytes: work,
        max_depth: 3,
        allocations: 16,
        ..Default::default()
    };
    Ok(TableStyleVariationPayload { bytes, report })
}

/// Decode one StylesheetArchive registry projection.
pub fn decode_stylesheet(
    source: &[u8],
    options: DecodeOptions,
) -> Result<StylesheetSnapshot<'_>, DecodeError> {
    Ok(decode_stylesheet_with_report(source, options)?.0)
}

/// Decode one StylesheetArchive registry projection with exact consumption.
pub fn decode_stylesheet_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(StylesheetSnapshot<'_>, DecodeReport), DecodeError> {
    ensure_input(source, options)?;
    let (snapshot, report) = scan_stylesheet(
        source,
        options.max_fields,
        options.max_work_bytes,
        options.recursion_limit,
        options.max_styles,
        None,
    )?;
    Ok((
        StylesheetSnapshot {
            source,
            style_count: snapshot.style_count,
            parent_identifier: snapshot.parent_identifier,
        },
        report,
    ))
}

/// Prepare appending a style ref, identified-style entry, and parent-child
/// edge. Existing registry bytes and field order remain untouched.
pub fn prepare_stylesheet_append<'source>(
    source: &'source [u8],
    append: StylesheetStyleAppend,
    options: DecodeOptions,
) -> Result<PreparedStylesheetAppend<'source>, DecodeError> {
    ensure_input(source, options)?;
    let (scan, report) = scan_stylesheet(
        source,
        options.max_fields,
        options.max_work_bytes,
        options.recursion_limit,
        options.max_styles,
        Some(append.style_identifier),
    )?;
    if scan.style_count.saturating_add(1) > options.max_styles {
        return Err(DecodeError::limit(DecodeLimit::Styles {
            observed: scan.style_count.saturating_add(1),
            maximum: options.max_styles,
        }));
    }
    if scan.style_ids[..scan.style_ids_len].contains(&append.style_identifier) {
        return Err(DecodeError::invalid("stylesheet already contains style"));
    }
    if let Some(parent) = append.parent_identifier {
        if !scan.style_ids[..scan.style_ids_len].contains(&parent) {
            return Err(DecodeError::invalid("stylesheet parent style is missing"));
        }
    }
    let parent_entry_exists = append.parent_identifier.is_some_and(|parent| {
        scan.child_entry_parents[..scan.child_entry_parent_len].contains(&parent)
    });
    let style_reference_len = reference_payload_len(append.style_identifier);
    let style_field_len = length_field_len(SHEET_STYLES_FIELD, style_reference_len);
    let parent_edge_len = append.parent_identifier.map_or(0, |parent| {
        let edge_payload_len =
            length_field_len(CHILDREN_PARENT_FIELD, reference_payload_len(parent))
                .saturating_add(length_field_len(CHILDREN_STYLE_FIELD, style_reference_len));
        if parent_entry_exists {
            length_field_len(CHILDREN_STYLE_FIELD, style_reference_len)
        } else {
            length_field_len(SHEET_CHILDREN_FIELD, edge_payload_len)
        }
    });
    let output_bytes = source
        .len()
        .saturating_add(style_field_len)
        .saturating_add(parent_edge_len);
    // One field plus one nested identifier for `styles`. A new parent edge
    // contributes its outer field, parent/child fields, and both nested
    // reference identifiers; extending an existing edge adds only the child
    // field and its nested identifier.
    let added_fields = 2 + match append.parent_identifier {
        None => 0,
        Some(_) if parent_entry_exists => 2,
        Some(_) => 5,
    };
    let requirements = RewriteExecutionRequirements {
        input_bytes: source.len(),
        output_bytes,
        fields: report
            .fields()
            .saturating_mul(2)
            .saturating_add(added_fields),
        work_bytes: report
            .work_bytes()
            .saturating_mul(2)
            .saturating_add(source.len())
            .saturating_add(output_bytes)
            .saturating_add(output_bytes)
            .saturating_add(style_reference_len)
            .saturating_add(parent_edge_len),
        max_depth: report.max_depth().max(3),
        allocations: 24,
    };
    validate_requirements_against_options(requirements, options)?;
    Ok(PreparedRewrite {
        source,
        kind: RewriteKind::StylesheetAppend {
            append,
            parent_entry_exists,
        },
        candidate_max_styles: options.max_styles,
        prepare_report: report,
        requirements,
    })
}

/// One-shot StylesheetArchive registry append.
pub fn append_stylesheet_style(
    source: &[u8],
    append: StylesheetStyleAppend,
    options: DecodeOptions,
) -> Result<(Vec<u8>, RewriteReport), DecodeError> {
    let prepared = prepare_stylesheet_append(source, append, options)?;
    let limits = prepared.execution_requirements().exact_limits();
    prepared.execute(limits)
}

fn ensure_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    if source.len() > options.max_input_bytes {
        return Err(DecodeError::limit(DecodeLimit::InputBytes {
            observed: source.len(),
            maximum: options.max_input_bytes,
        }));
    }
    if options.recursion_limit == 0 || options.recursion_limit > MAX_DEFAULT_RECURSION {
        return Err(DecodeError::limit(DecodeLimit::Nesting {
            observed: options.recursion_limit,
            maximum: MAX_DEFAULT_RECURSION,
        }));
    }
    Ok(())
}

fn validate_requirements_against_options(
    req: RewriteExecutionRequirements,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    if req.input_bytes > options.max_input_bytes {
        return Err(DecodeError::limit(DecodeLimit::InputBytes {
            observed: req.input_bytes,
            maximum: options.max_input_bytes,
        }));
    }
    if req.output_bytes > options.max_output_bytes {
        return Err(DecodeError::limit(DecodeLimit::OutputBytes {
            observed: req.output_bytes,
            maximum: options.max_output_bytes,
        }));
    }
    if req.fields > options.max_fields {
        return Err(DecodeError::limit(DecodeLimit::Fields {
            observed: req.fields,
            maximum: options.max_fields,
        }));
    }
    if req.work_bytes > options.max_work_bytes {
        return Err(DecodeError::limit(DecodeLimit::WorkBytes {
            observed: req.work_bytes,
            maximum: options.max_work_bytes,
        }));
    }
    if req.max_depth > options.recursion_limit {
        return Err(DecodeError::limit(DecodeLimit::Nesting {
            observed: req.max_depth,
            maximum: options.recursion_limit,
        }));
    }
    if req.allocations > options.max_allocations {
        return Err(DecodeError::limit(DecodeLimit::Allocations {
            observed: req.allocations,
            maximum: options.max_allocations,
        }));
    }
    Ok(())
}

fn validate_requirements(
    req: RewriteExecutionRequirements,
    limits: RewriteExecutionLimits,
) -> Result<(), DecodeError> {
    validate_requirements_against_options(
        req,
        DecodeOptions::new(
            limits.max_input_bytes,
            limits.max_output_bytes,
            limits.max_fields,
            limits.max_work_bytes,
            limits.max_depth,
            usize::MAX,
        )
        .with_max_allocations(limits.max_allocations),
    )
}

#[derive(Debug, Clone, Copy)]
struct Scan {
    style_identifier: u64,
    style_preset_identifier: Option<u64>,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    nested_ref_bytes: usize,
}

impl Scan {
    fn report(self, input_bytes: usize, output_bytes: usize) -> DecodeReport {
        DecodeReport {
            input_bytes,
            output_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            allocations: 0,
        }
    }
    fn nested_reference_bytes(self, new_len: usize) -> usize {
        self.nested_ref_bytes
            .saturating_sub(self.nested_ref_bytes.min(10))
            .saturating_add(new_len)
    }
}

fn scan_model(
    source: &[u8],
    max_fields: usize,
    max_work: usize,
    max_depth: u32,
) -> Result<Scan, DecodeError> {
    let options = DecodeOptions::new(
        source.len(),
        usize::MAX,
        max_fields,
        max_work,
        max_depth,
        usize::MAX,
    );
    let mut budget = Budget::new(options);
    let mut style = None;
    let mut preset = None;
    let mut style_ref_bytes = 0usize;
    scan_fields(source, 1, &mut budget, |field, budget| {
        match field.number {
            MODEL_STYLE_FIELD => {
                if style.is_some() {
                    return Err(DecodeError::invalid("duplicate TableModel table_style"));
                }
                let payload = field.length()?;
                let reference = parse_reference(payload, 2, budget)?;
                style_ref_bytes = payload.len();
                style = Some(reference);
            },
            MODEL_STYLE_PRESET_FIELD => {
                if preset.is_some() {
                    return Err(DecodeError::invalid(
                        "duplicate TableModel table_style_preset",
                    ));
                }
                let payload = field.length()?;
                preset = Some(parse_reference(payload, 2, budget)?);
            },
            _ => {},
        }
        Ok(())
    })?;
    let style_identifier =
        style.ok_or_else(|| DecodeError::invalid("missing TableModel table_style"))?;
    Ok(Scan {
        style_identifier,
        style_preset_identifier: preset,
        fields: budget.fields,
        work_bytes: budget.work,
        max_depth: budget.max_depth,
        nested_ref_bytes: style_ref_bytes,
    })
}

#[derive(Debug, Clone, Copy)]
struct StyleParsed<'source> {
    style_identifier: Option<&'source str>,
    parent_identifier: Option<u64>,
    stylesheet_identifier: Option<u64>,
    is_variation: bool,
    overrides: AppearanceOverrides,
}

fn parse_table_style<'source>(
    source: &'source [u8],
    depth: u32,
    budget: &mut Budget,
) -> Result<StyleParsed<'source>, DecodeError> {
    let mut super_payload = None;
    let mut properties_payload = None;
    let mut override_count = None;
    scan_fields(source, depth, budget, |field, _budget| {
        match field.number {
            STYLE_SUPER_FIELD => {
                if super_payload.is_some() {
                    return Err(DecodeError::invalid("duplicate TableStyle super"));
                }
                super_payload = Some(field.length()?);
            },
            STYLE_OVERRIDE_COUNT_FIELD => {
                if field.wire != 0 {
                    return Err(DecodeError::invalid("TableStyle override_count wire"));
                }
                if override_count.is_some() {
                    return Err(DecodeError::invalid("duplicate TableStyle override_count"));
                }
                override_count = Some(canonical_value(field)?);
            },
            STYLE_PROPERTIES_FIELD => {
                if properties_payload.is_some() {
                    return Err(DecodeError::invalid(
                        "duplicate TableStyle table_properties",
                    ));
                }
                properties_payload = Some(field.length()?);
            },
            _ => {},
        }
        Ok(())
    })?;
    let super_payload =
        super_payload.ok_or_else(|| DecodeError::invalid("missing TableStyle super"))?;
    let (style_identifier, parent_identifier, stylesheet_identifier, is_variation) =
        parse_style_super(super_payload, depth.saturating_add(1), budget)?;
    let overrides = properties_payload.map_or(Ok(AppearanceOverrides::default()), |payload| {
        parse_properties(payload, depth.saturating_add(1), budget)
    })?;
    Ok(StyleParsed {
        style_identifier,
        parent_identifier,
        stylesheet_identifier,
        is_variation,
        overrides,
    })
}

fn parse_style_super<'source>(
    source: &'source [u8],
    depth: u32,
    budget: &mut Budget,
) -> Result<(Option<&'source str>, Option<u64>, Option<u64>, bool), DecodeError> {
    let mut name = None;
    let mut identifier = None;
    let mut parent = None;
    let mut stylesheet = None;
    let mut variation = None;
    scan_fields(source, depth, budget, |field, budget| {
        match field.number {
            STYLE_NAME_FIELD => {
                if name.is_some() {
                    return Err(DecodeError::invalid("duplicate StyleArchive name"));
                }
                name = Some(
                    std::str::from_utf8(field.bytes()?)
                        .map_err(|_| DecodeError::invalid("invalid StyleArchive name"))?,
                );
            },
            STYLE_IDENTIFIER_FIELD => {
                if identifier.is_some() {
                    return Err(DecodeError::invalid(
                        "duplicate StyleArchive style_identifier",
                    ));
                }
                identifier =
                    Some(std::str::from_utf8(field.bytes()?).map_err(|_| {
                        DecodeError::invalid("invalid StyleArchive style identifier")
                    })?);
            },
            STYLE_PARENT_FIELD => {
                if parent.is_some() {
                    return Err(DecodeError::invalid("duplicate StyleArchive parent"));
                }
                parent = Some(parse_reference(
                    field.length()?,
                    depth.saturating_add(1),
                    budget,
                )?);
            },
            STYLE_VARIATION_FIELD => {
                if field.wire != 0 {
                    return Err(DecodeError::invalid("StyleArchive is_variation wire"));
                }
                if variation.is_some() {
                    return Err(DecodeError::invalid("duplicate StyleArchive is_variation"));
                }
                variation = Some(canonical_bool(field)?);
            },
            STYLE_STYLESHEET_FIELD => {
                if stylesheet.is_some() {
                    return Err(DecodeError::invalid("duplicate StyleArchive stylesheet"));
                }
                stylesheet = Some(parse_reference(
                    field.length()?,
                    depth.saturating_add(1),
                    budget,
                )?);
            },
            _ => {},
        }
        Ok(())
    })?;
    Ok((identifier, parent, stylesheet, variation.unwrap_or(false)))
}

fn parse_properties(
    source: &[u8],
    depth: u32,
    budget: &mut Budget,
) -> Result<AppearanceOverrides, DecodeError> {
    let mut current = [None; 7];
    let mut legacy = [None; 7];
    scan_fields(source, depth, budget, |field, _budget| {
        let (slot, is_legacy) = match field.number {
            PROPERTIES_BANDED_ROWS_FIELD => (0, false),
            PROPERTIES_AUTO_RESIZE_FIELD => (1, false),
            PROPERTIES_BODY_HORIZONTAL_FIELD => (2, false),
            PROPERTIES_BODY_VERTICAL_FIELD => (3, false),
            PROPERTIES_HEADER_COLUMN_FIELD => (4, false),
            PROPERTIES_LEGACY_HEADER_COLUMN_FIELD => (4, true),
            PROPERTIES_HEADER_ROW_FIELD => (5, false),
            PROPERTIES_LEGACY_HEADER_ROW_FIELD => (5, true),
            PROPERTIES_FOOTER_ROW_FIELD => (6, false),
            PROPERTIES_LEGACY_FOOTER_ROW_FIELD => (6, true),
            _ => return Ok(()),
        };
        let value = canonical_bool(field)?;
        let target = if is_legacy {
            &mut legacy[slot]
        } else {
            &mut current[slot]
        };
        if target.replace(value).is_some() {
            return Err(DecodeError::invalid("duplicate appearance override"));
        }
        Ok(())
    })?;
    let mut values = [None; 7];
    for slot in 0..7 {
        if let (Some(old), Some(new)) = (legacy[slot], current[slot]) {
            if old != new {
                return Err(DecodeError::invalid(
                    "conflicting legacy and modern appearance override",
                ));
            }
        }
        values[slot] = current[slot].or(legacy[slot]);
    }
    Ok(AppearanceOverrides {
        row_banding: values[0],
        row_sizing: values[1],
        body_horizontal: values[2],
        body_vertical: values[3],
        header_columns_horizontal: values[4],
        header_rows_vertical: values[5],
        footer_rows_vertical: values[6],
    })
}

fn merge_overrides(target: &mut AppearanceOverrides, source: AppearanceOverrides) {
    if target.row_banding.is_none() {
        target.row_banding = source.row_banding;
    }
    if target.row_sizing.is_none() {
        target.row_sizing = source.row_sizing;
    }
    if target.body_horizontal.is_none() {
        target.body_horizontal = source.body_horizontal;
    }
    if target.body_vertical.is_none() {
        target.body_vertical = source.body_vertical;
    }
    if target.header_columns_horizontal.is_none() {
        target.header_columns_horizontal = source.header_columns_horizontal;
    }
    if target.header_rows_vertical.is_none() {
        target.header_rows_vertical = source.header_rows_vertical;
    }
    if target.footer_rows_vertical.is_none() {
        target.footer_rows_vertical = source.footer_rows_vertical;
    }
}

#[derive(Debug, Clone, Copy)]
struct StylesheetScan {
    style_count: usize,
    style_ids: [u64; MAX_REGISTRY_FACTS],
    style_ids_len: usize,
    identified_name_hashes: [u64; MAX_REGISTRY_FACTS],
    identified_name_len: usize,
    identified_style_ids: [u64; MAX_REGISTRY_FACTS],
    identified_style_len: usize,
    child_edges: [(u64, u64); MAX_REGISTRY_FACTS],
    child_edges_len: usize,
    child_entry_parents: [u64; MAX_REGISTRY_FACTS],
    child_entry_parent_len: usize,
    parent_identifier: Option<u64>,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
}

impl StylesheetScan {
    fn report(self, source_len: usize) -> DecodeReport {
        DecodeReport {
            input_bytes: source_len,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            ..Default::default()
        }
    }
}

fn scan_stylesheet(
    source: &[u8],
    max_fields: usize,
    max_work: usize,
    max_depth: u32,
    max_styles: usize,
    reserved_style_identifier: Option<u64>,
) -> Result<(StylesheetScan, DecodeReport), DecodeError> {
    let options = DecodeOptions::new(
        source.len(),
        usize::MAX,
        max_fields,
        max_work,
        max_depth,
        max_styles,
    );
    let mut budget = Budget::new(options);
    let mut result = StylesheetScan {
        style_count: 0,
        style_ids: [0; MAX_REGISTRY_FACTS],
        style_ids_len: 0,
        identified_name_hashes: [0; MAX_REGISTRY_FACTS],
        identified_name_len: 0,
        identified_style_ids: [0; MAX_REGISTRY_FACTS],
        identified_style_len: 0,
        child_edges: [(0, 0); MAX_REGISTRY_FACTS],
        child_edges_len: 0,
        child_entry_parents: [0; MAX_REGISTRY_FACTS],
        child_entry_parent_len: 0,
        parent_identifier: None,
        fields: 0,
        work_bytes: 0,
        max_depth: 1,
    };
    let mut versioned_fields = [false; 16];
    scan_fields(source, 1, &mut budget, |field, budget| {
        match field.number {
            SHEET_STYLES_FIELD => {
                let id = parse_reference(field.length()?, 2, budget)?;
                if result.style_ids[..result.style_ids_len].contains(&id) {
                    return Err(DecodeError::invalid("duplicate stylesheet style"));
                }
                if result.style_count >= max_styles
                    || result.style_ids_len == result.style_ids.len()
                {
                    return Err(DecodeError::limit(DecodeLimit::Styles {
                        observed: result.style_count.saturating_add(1),
                        maximum: max_styles.min(MAX_REGISTRY_FACTS),
                    }));
                }
                result.style_ids[result.style_ids_len] = id;
                result.style_ids_len += 1;
                result.style_count += 1;
            },
            SHEET_IDENTIFIED_STYLES_FIELD => {
                let (name_hash, style_id) = parse_identified_entry(field.length()?, 2, budget)?;
                if result.identified_name_hashes[..result.identified_name_len].contains(&name_hash)
                {
                    return Err(DecodeError::invalid("duplicate identified style name"));
                }
                if result.identified_style_ids[..result.identified_style_len].contains(&style_id) {
                    return Err(DecodeError::invalid("duplicate identified style reference"));
                }
                if result.identified_name_len == MAX_REGISTRY_FACTS {
                    return Err(DecodeError::limit(DecodeLimit::Styles {
                        observed: result.identified_name_len.saturating_add(1),
                        maximum: max_styles.min(MAX_REGISTRY_FACTS),
                    }));
                }
                result.identified_name_hashes[result.identified_name_len] = name_hash;
                result.identified_style_ids[result.identified_style_len] = style_id;
                result.identified_name_len += 1;
                result.identified_style_len += 1;
            },
            SHEET_PARENT_FIELD => {
                if result.parent_identifier.is_some() {
                    return Err(DecodeError::invalid("duplicate stylesheet parent"));
                }
                result.parent_identifier = Some(parse_reference(field.length()?, 2, budget)?);
            },
            SHEET_CHILDREN_FIELD => {
                let (parent, children, children_len) =
                    parse_children_entry(field.length()?, 2, budget)?;
                if result.child_entry_parents[..result.child_entry_parent_len].contains(&parent) {
                    return Err(DecodeError::invalid(
                        "duplicate stylesheet parent-child entry",
                    ));
                }
                if result.child_entry_parent_len == MAX_REGISTRY_FACTS {
                    return Err(DecodeError::limit(DecodeLimit::Styles {
                        observed: result.child_entry_parent_len.saturating_add(1),
                        maximum: max_styles.min(MAX_REGISTRY_FACTS),
                    }));
                }
                result.child_entry_parents[result.child_entry_parent_len] = parent;
                result.child_entry_parent_len += 1;
                for child in children.into_iter().take(children_len) {
                    if result.child_edges[..result.child_edges_len].iter().any(
                        |(seen_parent, seen_child)| *seen_parent == parent && *seen_child == child,
                    ) {
                        return Err(DecodeError::invalid("duplicate stylesheet child edge"));
                    }
                    if result.child_edges[..result.child_edges_len].iter().any(
                        |(seen_parent, seen_child)| *seen_child == child && *seen_parent != parent,
                    ) {
                        return Err(DecodeError::invalid(
                            "stylesheet child has multiple parents",
                        ));
                    }
                    if result.child_edges_len == MAX_REGISTRY_FACTS {
                        return Err(DecodeError::limit(DecodeLimit::Styles {
                            observed: result.child_edges_len.saturating_add(1),
                            maximum: max_styles.min(MAX_REGISTRY_FACTS),
                        }));
                    }
                    result.child_edges[result.child_edges_len] = (parent, child);
                    result.child_edges_len += 1;
                }
            },
            SHEET_VERSIONED_FIRST_FIELD..=SHEET_VERSIONED_LAST_FIELD => {
                let index = usize::try_from(field.number - SHEET_VERSIONED_FIRST_FIELD)
                    .map_err(|_| DecodeError::invalid("versioned stylesheet field overflow"))?;
                if versioned_fields[index] {
                    return Err(DecodeError::invalid(
                        "duplicate versioned stylesheet registry",
                    ));
                }
                versioned_fields[index] = true;
                validate_versioned_styles(
                    field.length()?,
                    2,
                    budget,
                    max_styles,
                    reserved_style_identifier,
                )?;
            },
            _ => {},
        }
        Ok(())
    })?;
    for style_id in result.identified_style_ids[..result.identified_style_len].iter() {
        if !result.style_ids[..result.style_ids_len].contains(style_id) {
            return Err(DecodeError::invalid(
                "identified style is absent from stylesheet styles",
            ));
        }
    }
    for (parent, child) in result.child_edges[..result.child_edges_len].iter() {
        if !result.style_ids[..result.style_ids_len].contains(parent)
            || !result.style_ids[..result.style_ids_len].contains(child)
        {
            return Err(DecodeError::invalid(
                "stylesheet child edge references an unregistered style",
            ));
        }
    }
    if let Some(parent) = result.parent_identifier {
        if !result.style_ids[..result.style_ids_len].contains(&parent) {
            return Err(DecodeError::invalid(
                "stylesheet parent is absent from stylesheet styles",
            ));
        }
    }
    result.fields = budget.fields;
    result.work_bytes = budget.work;
    result.max_depth = budget.max_depth;
    Ok((result, result.report(source.len())))
}

fn validate_versioned_styles(
    source: &[u8],
    depth: u32,
    budget: &mut Budget,
    max_styles: usize,
    reserved_style_identifier: Option<u64>,
) -> Result<(), DecodeError> {
    let maximum = max_styles.min(MAX_REGISTRY_FACTS);
    let mut style_ids = [0u64; MAX_REGISTRY_FACTS];
    let mut style_ids_len = 0usize;
    let mut identified_name_hashes = [0u64; MAX_REGISTRY_FACTS];
    let mut identified_style_ids = [0u64; MAX_REGISTRY_FACTS];
    let mut identified_len = 0usize;
    let mut child_entry_parents = [0u64; MAX_REGISTRY_FACTS];
    let mut child_entry_parent_len = 0usize;
    let mut child_edges = [(0u64, 0u64); MAX_REGISTRY_FACTS];
    let mut child_edges_len = 0usize;

    scan_fields(source, depth, budget, |field, budget| {
        match field.number {
            VERSIONED_STYLES_FIELD => {
                let identifier = parse_reference(field.length()?, depth.saturating_add(1), budget)?;
                if reserved_style_identifier == Some(identifier) {
                    return Err(DecodeError::invalid(
                        "style identifier is reserved by a versioned registry",
                    ));
                }
                if style_ids[..style_ids_len].contains(&identifier) {
                    return Err(DecodeError::invalid("duplicate versioned stylesheet style"));
                }
                if style_ids_len >= maximum {
                    return Err(DecodeError::limit(DecodeLimit::Styles {
                        observed: style_ids_len.saturating_add(1),
                        maximum,
                    }));
                }
                style_ids[style_ids_len] = identifier;
                style_ids_len += 1;
            },
            VERSIONED_IDENTIFIED_STYLES_FIELD => {
                let (name_hash, style_id) =
                    parse_identified_entry(field.length()?, depth.saturating_add(1), budget)?;
                if identified_name_hashes[..identified_len].contains(&name_hash) {
                    return Err(DecodeError::invalid(
                        "duplicate versioned identified style name",
                    ));
                }
                if identified_style_ids[..identified_len].contains(&style_id) {
                    return Err(DecodeError::invalid(
                        "duplicate versioned identified style reference",
                    ));
                }
                if identified_len >= maximum {
                    return Err(DecodeError::limit(DecodeLimit::Styles {
                        observed: identified_len.saturating_add(1),
                        maximum,
                    }));
                }
                identified_name_hashes[identified_len] = name_hash;
                identified_style_ids[identified_len] = style_id;
                identified_len += 1;
            },
            VERSIONED_CHILDREN_FIELD => {
                let (parent, children, children_len) =
                    parse_children_entry(field.length()?, depth.saturating_add(1), budget)?;
                if child_entry_parents[..child_entry_parent_len].contains(&parent) {
                    return Err(DecodeError::invalid(
                        "duplicate versioned stylesheet parent-child entry",
                    ));
                }
                if child_entry_parent_len >= maximum {
                    return Err(DecodeError::limit(DecodeLimit::Styles {
                        observed: child_entry_parent_len.saturating_add(1),
                        maximum,
                    }));
                }
                child_entry_parents[child_entry_parent_len] = parent;
                child_entry_parent_len += 1;
                for child in children.into_iter().take(children_len) {
                    if child_edges[..child_edges_len]
                        .iter()
                        .any(|(seen_parent, seen_child)| {
                            *seen_parent == parent && *seen_child == child
                        })
                    {
                        return Err(DecodeError::invalid(
                            "duplicate versioned stylesheet child edge",
                        ));
                    }
                    if child_edges[..child_edges_len]
                        .iter()
                        .any(|(seen_parent, seen_child)| {
                            *seen_child == child && *seen_parent != parent
                        })
                    {
                        return Err(DecodeError::invalid(
                            "versioned stylesheet child has multiple parents",
                        ));
                    }
                    if child_edges_len >= maximum {
                        return Err(DecodeError::limit(DecodeLimit::Styles {
                            observed: child_edges_len.saturating_add(1),
                            maximum,
                        }));
                    }
                    child_edges[child_edges_len] = (parent, child);
                    child_edges_len += 1;
                }
            },
            _ => {},
        }
        Ok(())
    })?;

    for style_id in &identified_style_ids[..identified_len] {
        if !style_ids[..style_ids_len].contains(style_id) {
            return Err(DecodeError::invalid(
                "versioned identified style is absent from versioned styles",
            ));
        }
    }
    for (parent, child) in &child_edges[..child_edges_len] {
        if !style_ids[..style_ids_len].contains(parent)
            || !style_ids[..style_ids_len].contains(child)
        {
            return Err(DecodeError::invalid(
                "versioned child edge references an unregistered style",
            ));
        }
    }
    Ok(())
}

fn parse_identified_entry(
    source: &[u8],
    depth: u32,
    budget: &mut Budget,
) -> Result<(u64, u64), DecodeError> {
    let mut identifier_hash = None;
    let mut style_id = None;
    scan_fields(source, depth, budget, |field, budget| {
        match field.number {
            IDENTIFIED_IDENTIFIER_FIELD => {
                if identifier_hash.is_some() {
                    return Err(DecodeError::invalid(
                        "duplicate identified style identifier",
                    ));
                }
                let identifier = std::str::from_utf8(field.bytes()?)
                    .map_err(|_| DecodeError::invalid("invalid identified style identifier"))?;
                identifier_hash = Some(hash_identifier(identifier.as_bytes()));
            },
            IDENTIFIED_STYLE_FIELD => {
                if style_id.is_some() {
                    return Err(DecodeError::invalid("duplicate identified style reference"));
                }
                style_id = Some(parse_reference(
                    field.length()?,
                    depth.saturating_add(1),
                    budget,
                )?);
            },
            _ => {},
        }
        Ok(())
    })?;
    let (Some(identifier_hash), Some(style_id)) = (identifier_hash, style_id) else {
        return Err(DecodeError::invalid("identified style entry is incomplete"));
    };
    Ok((identifier_hash, style_id))
}

fn parse_children_entry(
    source: &[u8],
    depth: u32,
    budget: &mut Budget,
) -> Result<(u64, [u64; MAX_REGISTRY_FACTS], usize), DecodeError> {
    let mut parent = false;
    let mut parent_identifier = 0;
    let mut children = [0u64; MAX_REGISTRY_FACTS];
    let mut children_len = 0usize;
    scan_fields(source, depth, budget, |field, budget| {
        match field.number {
            CHILDREN_PARENT_FIELD => {
                if parent {
                    return Err(DecodeError::invalid("duplicate style children parent"));
                }
                parent = true;
                parent_identifier =
                    parse_reference(field.length()?, depth.saturating_add(1), budget)?;
            },
            CHILDREN_STYLE_FIELD => {
                if children_len == MAX_REGISTRY_FACTS {
                    return Err(DecodeError::limit(DecodeLimit::Styles {
                        observed: children_len.saturating_add(1),
                        maximum: budget.options.max_styles.min(MAX_REGISTRY_FACTS),
                    }));
                }
                children[children_len] =
                    parse_reference(field.length()?, depth.saturating_add(1), budget)?;
                if children[..children_len].contains(&children[children_len]) {
                    return Err(DecodeError::invalid("duplicate style child reference"));
                }
                children_len += 1;
            },
            _ => {},
        }
        Ok(())
    })?;
    if !parent {
        return Err(DecodeError::invalid("style children entry has no parent"));
    }
    Ok((parent_identifier, children, children_len))
}

fn parse_reference(source: &[u8], depth: u32, budget: &mut Budget) -> Result<u64, DecodeError> {
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut deprecated_is_external = None;
    scan_fields(source, depth, budget, |field, _budget| {
        match field.number {
            MODEL_STYLE_REFERENCE_FIELD => {
                if identifier.is_some() {
                    return Err(DecodeError::invalid("duplicate reference identifier"));
                }
                identifier = Some(canonical_value(field)?);
            },
            2 => {
                if field.wire != 0 {
                    return Err(DecodeError::invalid("reference type wire"));
                }
                if deprecated_type.is_some() {
                    return Err(DecodeError::invalid("duplicate reference type"));
                }
                let value = canonical_value(field)?;
                if value != 0 {
                    return Err(DecodeError::invalid("unsupported reference type"));
                }
                deprecated_type = Some(value);
            },
            3 => {
                if deprecated_is_external.is_some() {
                    return Err(DecodeError::invalid("duplicate reference external flag"));
                }
                if canonical_bool(field)? {
                    return Err(DecodeError::invalid("external appearance reference"));
                }
                deprecated_is_external = Some(false);
            },
            _ => {},
        }
        Ok(())
    })?;
    identifier
        .filter(|value| *value != 0)
        .ok_or_else(|| DecodeError::invalid("reference identifier is missing or zero"))
}

struct Budget {
    options: DecodeOptions,
    fields: usize,
    work: usize,
    max_depth: u32,
}

impl Budget {
    fn new(options: DecodeOptions) -> Self {
        Self {
            options,
            fields: 0,
            work: 0,
            max_depth: 0,
        }
    }
    fn charge(&mut self, bytes: usize, depth: u32) -> Result<(), DecodeError> {
        self.fields = self
            .fields
            .checked_add(1)
            .ok_or_else(|| DecodeError::invalid("field count overflow"))?;
        self.work = self
            .work
            .checked_add(bytes)
            .ok_or_else(|| DecodeError::invalid("work count overflow"))?;
        self.max_depth = self.max_depth.max(depth);
        if self.fields > self.options.max_fields {
            return Err(DecodeError::limit(DecodeLimit::Fields {
                observed: self.fields,
                maximum: self.options.max_fields,
            }));
        }
        if self.work > self.options.max_work_bytes {
            return Err(DecodeError::limit(DecodeLimit::WorkBytes {
                observed: self.work,
                maximum: self.options.max_work_bytes,
            }));
        }
        if depth > self.options.recursion_limit {
            return Err(DecodeError::limit(DecodeLimit::Nesting {
                observed: depth,
                maximum: self.options.recursion_limit,
            }));
        }
        Ok(())
    }
    fn charge_fields(&mut self, fields: usize, depth: u32) -> Result<(), DecodeError> {
        for _ in 0..fields {
            self.charge(1, depth)?;
        }
        Ok(())
    }
    fn report(&self, input: usize, output: usize) -> DecodeReport {
        DecodeReport {
            input_bytes: input,
            output_bytes: output,
            fields: self.fields,
            work_bytes: self.work,
            max_depth: self.max_depth,
            allocations: 0,
        }
    }
}

#[derive(Debug, Clone, Copy)]
struct Field<'source> {
    number: u32,
    wire: u8,
    raw: &'source [u8],
    payload: &'source [u8],
}

impl<'source> Field<'source> {
    fn length(self) -> Result<&'source [u8], DecodeError> {
        if self.wire == 2 {
            Ok(self.payload)
        } else {
            Err(DecodeError::invalid("expected length-delimited field"))
        }
    }
    fn bytes(self) -> Result<&'source [u8], DecodeError> {
        self.length()
    }
}

fn scan_fields<'source, F>(
    source: &'source [u8],
    depth: u32,
    budget: &mut Budget,
    mut visit: F,
) -> Result<(), DecodeError>
where
    F: FnMut(Field<'source>, &mut Budget) -> Result<(), DecodeError>,
{
    if depth > budget.options.recursion_limit {
        return Err(DecodeError::limit(DecodeLimit::Nesting {
            observed: depth,
            maximum: budget.options.recursion_limit,
        }));
    }
    let mut offset = 0usize;
    while offset < source.len() {
        let (field, next) = parse_field(source, offset)?;
        if field.wire == 4 {
            return Err(DecodeError::invalid("unexpected end group"));
        }
        budget.charge(field.raw.len(), depth)?;
        visit(field, budget)?;
        if field.wire == 3 {
            scan_group_fields(field.payload, depth.saturating_add(1), budget)?;
        }
        offset = next;
    }
    Ok(())
}

fn scan_group_fields(source: &[u8], depth: u32, budget: &mut Budget) -> Result<(), DecodeError> {
    let mut offset = 0usize;
    while offset < source.len() {
        let (field, next) = parse_field(source, offset)?;
        budget.charge(field.raw.len(), depth)?;
        if field.wire == 4 {
            return Err(DecodeError::invalid("unexpected end group"));
        }
        if field.wire == 3 {
            scan_group_fields(field.payload, depth.saturating_add(1), budget)?;
        }
        offset = next;
    }
    Ok(())
}

fn parse_field(source: &[u8], start: usize) -> Result<(Field<'_>, usize), DecodeError> {
    let (key, key_len) = decode_varint(&source[start..], true)?;
    if key == 0 {
        return Err(DecodeError::invalid("zero protobuf key"));
    }
    let number =
        u32::try_from(key >> 3).map_err(|_| DecodeError::invalid("field number overflow"))?;
    if number == 0 || number > MAX_FIELD_NUMBER {
        return Err(DecodeError::invalid("field number out of range"));
    }
    let wire = u8::try_from(key & 7).map_err(|_| DecodeError::invalid("wire type overflow"))?;
    if wire == 6 || wire == 7 {
        return Err(DecodeError::invalid("reserved wire type"));
    }
    let cursor = start
        .checked_add(key_len)
        .ok_or_else(|| DecodeError::invalid("field overflow"))?;
    match wire {
        0 => {
            let (_, value_len) = decode_varint(&source[cursor..], false)?;
            let end = cursor
                .checked_add(value_len)
                .ok_or_else(|| DecodeError::invalid("field overflow"))?;
            if end > source.len() {
                return Err(DecodeError::invalid("truncated varint"));
            }
            Ok((
                Field {
                    number,
                    wire,
                    raw: &source[start..end],
                    payload: &source[cursor..end],
                },
                end,
            ))
        },
        1 => {
            let end = cursor
                .checked_add(8)
                .ok_or_else(|| DecodeError::invalid("field overflow"))?;
            if end > source.len() {
                return Err(DecodeError::invalid("truncated fixed64"));
            }
            Ok((
                Field {
                    number,
                    wire,
                    raw: &source[start..end],
                    payload: &source[cursor..end],
                },
                end,
            ))
        },
        2 => {
            let (length, length_len) = decode_varint(&source[cursor..], true)?;
            let payload_start = cursor
                .checked_add(length_len)
                .ok_or_else(|| DecodeError::invalid("field overflow"))?;
            let payload_end = payload_start
                .checked_add(
                    usize::try_from(length).map_err(|_| DecodeError::invalid("length overflow"))?,
                )
                .ok_or_else(|| DecodeError::invalid("field overflow"))?;
            if payload_end > source.len() {
                return Err(DecodeError::invalid("truncated length-delimited field"));
            }
            Ok((
                Field {
                    number,
                    wire,
                    raw: &source[start..payload_end],
                    payload: &source[payload_start..payload_end],
                },
                payload_end,
            ))
        },
        3 => {
            let payload_start = cursor;
            let (end, payload_end) = find_group_end(source, payload_start, number)?;
            Ok((
                Field {
                    number,
                    wire,
                    raw: &source[start..end],
                    payload: &source[payload_start..payload_end],
                },
                end,
            ))
        },
        4 => Ok((
            Field {
                number,
                wire,
                raw: &source[start..cursor],
                payload: &source[cursor..cursor],
            },
            cursor,
        )),
        5 => {
            let end = cursor
                .checked_add(4)
                .ok_or_else(|| DecodeError::invalid("field overflow"))?;
            if end > source.len() {
                return Err(DecodeError::invalid("truncated fixed32"));
            }
            Ok((
                Field {
                    number,
                    wire,
                    raw: &source[start..end],
                    payload: &source[cursor..end],
                },
                end,
            ))
        },
        _ => Err(DecodeError::invalid("unsupported wire type")),
    }
}

fn find_group_end(
    source: &[u8],
    mut offset: usize,
    group: u32,
) -> Result<(usize, usize), DecodeError> {
    let mut stack = [0u32; MAX_DEFAULT_RECURSION as usize];
    let mut stack_len = 1usize;
    stack[0] = group;
    while offset < source.len() {
        let field_start = offset;
        let (number, wire, next) = parse_field_shallow(source, offset)?;
        match wire {
            3 => {
                if stack_len == stack.len() {
                    return Err(DecodeError::limit(DecodeLimit::Nesting {
                        observed: stack_len.saturating_add(1) as u32,
                        maximum: MAX_DEFAULT_RECURSION,
                    }));
                }
                stack[stack_len] = number;
                stack_len += 1;
                offset = next;
            },
            4 => {
                if stack_len == 0 || stack[stack_len - 1] != number {
                    return Err(DecodeError::invalid("mismatched end group"));
                }
                stack_len -= 1;
                if stack_len == 0 {
                    return Ok((next, field_start));
                }
                offset = next;
            },
            _ => {
                offset = next;
            },
        }
    }
    Err(DecodeError::invalid("unterminated group"))
}

fn parse_field_shallow(source: &[u8], start: usize) -> Result<(u32, u8, usize), DecodeError> {
    let (key, key_len) = decode_varint(&source[start..], true)?;
    if key == 0 {
        return Err(DecodeError::invalid("zero protobuf key"));
    }
    let number =
        u32::try_from(key >> 3).map_err(|_| DecodeError::invalid("field number overflow"))?;
    if number == 0 || number > MAX_FIELD_NUMBER {
        return Err(DecodeError::invalid("field number out of range"));
    }
    let wire = u8::try_from(key & 7).map_err(|_| DecodeError::invalid("wire type overflow"))?;
    if wire == 6 || wire == 7 {
        return Err(DecodeError::invalid("reserved wire type"));
    }
    let cursor = start
        .checked_add(key_len)
        .ok_or_else(|| DecodeError::invalid("field overflow"))?;
    let end = match wire {
        0 => {
            let (_, value_len) = decode_varint(&source[cursor..], false)?;
            cursor
                .checked_add(value_len)
                .ok_or_else(|| DecodeError::invalid("field overflow"))?
        },
        1 => cursor
            .checked_add(8)
            .ok_or_else(|| DecodeError::invalid("field overflow"))?,
        2 => {
            let (length, length_len) = decode_varint(&source[cursor..], true)?;
            let payload_start = cursor
                .checked_add(length_len)
                .ok_or_else(|| DecodeError::invalid("field overflow"))?;
            payload_start
                .checked_add(
                    usize::try_from(length).map_err(|_| DecodeError::invalid("length overflow"))?,
                )
                .ok_or_else(|| DecodeError::invalid("field overflow"))?
        },
        3 | 4 => cursor,
        5 => cursor
            .checked_add(4)
            .ok_or_else(|| DecodeError::invalid("field overflow"))?,
        _ => return Err(DecodeError::invalid("unsupported wire type")),
    };
    if end > source.len() {
        return Err(DecodeError::invalid("truncated field"));
    }
    Ok((number, wire, end))
}

fn decode_varint(source: &[u8], canonical: bool) -> Result<(u64, usize), DecodeError> {
    let mut value = 0u64;
    for (index, byte) in source.iter().copied().take(10).enumerate() {
        if index == 9 && byte > 1 {
            return Err(DecodeError::invalid("varint overflow"));
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            let length = index + 1;
            if canonical && encoded_varint_len(value) != length {
                return Err(DecodeError::invalid("non-canonical varint"));
            }
            return Ok((value, length));
        }
    }
    Err(DecodeError::invalid("unterminated varint"))
}

fn canonical_value(field: Field<'_>) -> Result<u64, DecodeError> {
    if field.wire != 0 {
        return Err(DecodeError::invalid("expected varint field"));
    }
    let (value, length) = decode_varint(field.payload, true)?;
    if length != field.payload.len() {
        return Err(DecodeError::invalid("non-canonical known scalar"));
    }
    Ok(value)
}
fn canonical_bool(field: Field<'_>) -> Result<bool, DecodeError> {
    match canonical_value(field)? {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(DecodeError::invalid("boolean is not zero or one")),
    }
}

fn encoded_varint_len(mut value: u64) -> usize {
    let mut length = 1usize;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}
fn encode_varint(mut value: u64, output: &mut Vec<u8>) {
    while value >= 0x80 {
        output.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}
#[cfg(test)]
fn varint_field(number: u32, value: u64) -> Vec<u8> {
    let mut result = Vec::new();
    encode_varint(u64::from(number) << 3, &mut result);
    encode_varint(value, &mut result);
    result
}
fn varint_field_fallible(number: u32, value: u64) -> Result<Vec<u8>, DecodeError> {
    let capacity = varint_field_len(number, value);
    let mut result = Vec::new();
    result
        .try_reserve_exact(capacity)
        .map_err(|_| DecodeError::allocation(capacity))?;
    encode_varint(u64::from(number) << 3, &mut result);
    encode_varint(value, &mut result);
    Ok(result)
}
fn varint_field_len(number: u32, value: u64) -> usize {
    encoded_varint_len(u64::from(number) << 3) + encoded_varint_len(value)
}
#[cfg(test)]
fn bool_field(number: u32, value: bool) -> Vec<u8> {
    varint_field(number, u64::from(value))
}
fn bool_field_fallible(number: u32, value: bool) -> Result<Vec<u8>, DecodeError> {
    varint_field_fallible(number, u64::from(value))
}
#[cfg(test)]
fn length_field(number: u32, payload: &[u8]) -> Vec<u8> {
    let mut result = Vec::new();
    encode_varint((u64::from(number) << 3) | 2, &mut result);
    encode_varint(payload.len() as u64, &mut result);
    result.extend_from_slice(payload);
    result
}
fn length_field_fallible(number: u32, payload: &[u8]) -> Result<Vec<u8>, DecodeError> {
    let capacity = length_field_len(number, payload.len());
    let mut result = Vec::new();
    result
        .try_reserve_exact(capacity)
        .map_err(|_| DecodeError::allocation(capacity))?;
    encode_varint((u64::from(number) << 3) | 2, &mut result);
    encode_varint(payload.len() as u64, &mut result);
    result.extend_from_slice(payload);
    Ok(result)
}
fn length_field_len(number: u32, payload_len: usize) -> usize {
    encoded_varint_len((u64::from(number) << 3) | 2)
        + encoded_varint_len(payload_len as u64)
        + payload_len
}
fn reference_payload_len(identifier: u64) -> usize {
    varint_field_len(MODEL_STYLE_REFERENCE_FIELD, identifier)
}

fn hash_identifier(bytes: &[u8]) -> u64 {
    // FNV-1a is deterministic and allocation-free.  A collision is treated
    // as a duplicate, which is conservative for a hostile registry.
    bytes.iter().fold(0xcbf2_9ce4_8422_2325, |hash, byte| {
        (hash ^ u64::from(*byte)).wrapping_mul(0x1000_0000_01b3)
    })
}
#[cfg(test)]
fn reference_payload(identifier: u64) -> Vec<u8> {
    varint_field(MODEL_STYLE_REFERENCE_FIELD, identifier)
}
fn reference_payload_fallible(identifier: u64) -> Result<Vec<u8>, DecodeError> {
    varint_field_fallible(MODEL_STYLE_REFERENCE_FIELD, identifier)
}
#[cfg(test)]
fn canonical_properties(overrides: AppearanceOverrides) -> Vec<u8> {
    let mut result = Vec::new();
    for (field, value) in [
        (PROPERTIES_BANDED_ROWS_FIELD, overrides.row_banding),
        (PROPERTIES_AUTO_RESIZE_FIELD, overrides.row_sizing),
        (PROPERTIES_BODY_HORIZONTAL_FIELD, overrides.body_horizontal),
        (PROPERTIES_BODY_VERTICAL_FIELD, overrides.body_vertical),
        (
            PROPERTIES_HEADER_COLUMN_FIELD,
            overrides.header_columns_horizontal,
        ),
        (PROPERTIES_HEADER_ROW_FIELD, overrides.header_rows_vertical),
        (PROPERTIES_FOOTER_ROW_FIELD, overrides.footer_rows_vertical),
    ] {
        if let Some(value) = value {
            result.extend_from_slice(&bool_field(field, value));
        }
    }
    result
}

fn canonical_properties_fallible(overrides: AppearanceOverrides) -> Result<Vec<u8>, DecodeError> {
    let capacity = canonical_properties_len(overrides);
    let mut result = Vec::new();
    result
        .try_reserve_exact(capacity)
        .map_err(|_| DecodeError::allocation(capacity))?;
    for (field, value) in [
        (PROPERTIES_BANDED_ROWS_FIELD, overrides.row_banding),
        (PROPERTIES_AUTO_RESIZE_FIELD, overrides.row_sizing),
        (PROPERTIES_BODY_HORIZONTAL_FIELD, overrides.body_horizontal),
        (PROPERTIES_BODY_VERTICAL_FIELD, overrides.body_vertical),
        (
            PROPERTIES_HEADER_COLUMN_FIELD,
            overrides.header_columns_horizontal,
        ),
        (PROPERTIES_HEADER_ROW_FIELD, overrides.header_rows_vertical),
        (PROPERTIES_FOOTER_ROW_FIELD, overrides.footer_rows_vertical),
    ] {
        if let Some(value) = value {
            result.extend_from_slice(&bool_field_fallible(field, value)?);
        }
    }
    Ok(result)
}

fn canonical_properties_len(overrides: AppearanceOverrides) -> usize {
    [
        (PROPERTIES_BANDED_ROWS_FIELD, overrides.row_banding),
        (PROPERTIES_AUTO_RESIZE_FIELD, overrides.row_sizing),
        (PROPERTIES_BODY_HORIZONTAL_FIELD, overrides.body_horizontal),
        (PROPERTIES_BODY_VERTICAL_FIELD, overrides.body_vertical),
        (
            PROPERTIES_HEADER_COLUMN_FIELD,
            overrides.header_columns_horizontal,
        ),
        (PROPERTIES_HEADER_ROW_FIELD, overrides.header_rows_vertical),
        (PROPERTIES_FOOTER_ROW_FIELD, overrides.footer_rows_vertical),
    ]
    .into_iter()
    .map(|(field, value)| value.map_or(0, |value| varint_field_len(field, u64::from(value))))
    .sum()
}

fn complete_overrides(overrides: AppearanceOverrides) -> bool {
    overrides.row_banding.is_some()
        && overrides.row_sizing.is_some()
        && overrides.body_horizontal.is_some()
        && overrides.body_vertical.is_some()
        && overrides.header_columns_horizontal.is_some()
        && overrides.header_rows_vertical.is_some()
        && overrides.footer_rows_vertical.is_some()
}

fn rewrite_model_into(
    source: &[u8],
    old: u64,
    new: u64,
    output: &mut [u8],
) -> Result<(), DecodeError> {
    let mut position = 0usize;
    let mut written = 0usize;
    let mut replaced = false;
    while position < source.len() {
        let (field, next) = parse_field(source, position)?;
        if field.number == MODEL_STYLE_FIELD {
            let payload = field.length()?;
            let nested_capacity = payload.len().saturating_add(10);
            let mut nested = Vec::new();
            nested
                .try_reserve_exact(nested_capacity)
                .map_err(|_| DecodeError::allocation(nested_capacity))?;
            let mut nested_pos = 0usize;
            while nested_pos < payload.len() {
                let (nested_field, nested_next) = parse_field(payload, nested_pos)?;
                if nested_field.number == MODEL_STYLE_REFERENCE_FIELD {
                    if replaced {
                        return Err(DecodeError::invalid("duplicate model style reference"));
                    }
                    let value = canonical_value(nested_field)?;
                    if value != old {
                        return Err(DecodeError::invalid(
                            "model style changed between prepare and execute",
                        ));
                    }
                    nested.extend_from_slice(&varint_field_fallible(
                        MODEL_STYLE_REFERENCE_FIELD,
                        new,
                    )?);
                    replaced = true;
                } else {
                    nested.extend_from_slice(nested_field.raw);
                }
                nested_pos = nested_next;
            }
            let field_bytes = length_field_fallible(MODEL_STYLE_FIELD, &nested)?;
            output[written..written + field_bytes.len()].copy_from_slice(&field_bytes);
            written += field_bytes.len();
        } else {
            output[written..written + field.raw.len()].copy_from_slice(field.raw);
            written += field.raw.len();
        }
        position = next;
    }
    if !replaced || written != output.len() {
        return Err(DecodeError::invalid(
            "model style rewrite did not replace exactly one edge",
        ));
    }
    Ok(())
}

#[cfg(test)]
fn children_entry(parent: u64, child: u64) -> Vec<u8> {
    let mut nested = Vec::new();
    nested.extend_from_slice(&length_field(
        CHILDREN_PARENT_FIELD,
        &reference_payload(parent),
    ));
    nested.extend_from_slice(&length_field(
        CHILDREN_STYLE_FIELD,
        &reference_payload(child),
    ));
    nested
}

fn children_entry_fallible(parent: u64, child: u64) -> Result<Vec<u8>, DecodeError> {
    let parent_reference = reference_payload_fallible(parent)?;
    let child_reference = reference_payload_fallible(child)?;
    let parent_field = length_field_fallible(CHILDREN_PARENT_FIELD, &parent_reference)?;
    let child_field = length_field_fallible(CHILDREN_STYLE_FIELD, &child_reference)?;
    let capacity = parent_field.len().saturating_add(child_field.len());
    let mut nested = Vec::new();
    nested
        .try_reserve_exact(capacity)
        .map_err(|_| DecodeError::allocation(capacity))?;
    nested.extend_from_slice(&parent_field);
    nested.extend_from_slice(&child_field);
    Ok(nested)
}

fn append_stylesheet_into(
    source: &[u8],
    append: StylesheetStyleAppend,
    parent_entry_exists: bool,
    output: &mut [u8],
) -> Result<(), DecodeError> {
    // The source registry was already fully scanned and charged during
    // prepare.  This small lookup must not re-run it under the caller's
    // residual field/work ceilings (which would make exact prepared replay
    // depend on an unreported second limit).  Keep the traversal bounded by
    // the immutable source length and the codec's hard nesting/registry
    // caps; the prepared requirements own its aggregate cost.
    let lookup_options = DecodeOptions::new(
        source.len(),
        usize::MAX,
        usize::MAX,
        usize::MAX,
        MAX_DEFAULT_RECURSION,
        MAX_REGISTRY_FACTS,
    );
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(output.len())
        .map_err(|_| DecodeError::allocation(output.len()))?;
    let mut offset = 0usize;
    let mut replaced_parent = false;
    while offset < source.len() {
        let (field, next) = parse_field(source, offset)?;
        if field.number == SHEET_CHILDREN_FIELD
            && parent_entry_exists
            && append.parent_identifier.is_some()
        {
            let (parent, _, _) =
                parse_children_entry(field.length()?, 2, &mut Budget::new(lookup_options))?;
            if Some(parent) == append.parent_identifier {
                let child_reference = reference_payload_fallible(append.style_identifier)?;
                let child = length_field_fallible(CHILDREN_STYLE_FIELD, &child_reference)?;
                let edge_capacity = field.payload.len().saturating_add(child.len());
                let mut edge = Vec::new();
                edge.try_reserve_exact(edge_capacity)
                    .map_err(|_| DecodeError::allocation(edge_capacity))?;
                edge.extend_from_slice(field.payload);
                edge.extend_from_slice(&child);
                bytes.extend_from_slice(&length_field_fallible(SHEET_CHILDREN_FIELD, &edge)?);
                replaced_parent = true;
            } else {
                bytes.extend_from_slice(field.raw);
            }
        } else {
            bytes.extend_from_slice(field.raw);
        }
        offset = next;
    }
    let style_reference = reference_payload_fallible(append.style_identifier)?;
    bytes.extend_from_slice(&length_field_fallible(
        SHEET_STYLES_FIELD,
        &style_reference,
    )?);
    if append.parent_identifier.is_some() && !replaced_parent {
        let Some(parent) = append.parent_identifier else {
            return Err(DecodeError::invalid("missing stylesheet parent"));
        };
        let edge = children_entry_fallible(parent, append.style_identifier)?;
        bytes.extend_from_slice(&length_field_fallible(SHEET_CHILDREN_FIELD, &edge)?);
    }
    if bytes.len() != output.len() {
        return Err(DecodeError::invalid(
            "stylesheet output size changed during execute",
        ));
    }
    output.copy_from_slice(&bytes);
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::for_source(source).with_max_output_bytes(source.len().saturating_mul(4))
    }
    fn model(style: u64, unknown: bool) -> Vec<u8> {
        let mut nested = reference_payload(style);
        if unknown {
            nested.extend_from_slice(&[0x98, 0x03, 0x81, 0x80, 0x00]);
        }
        let mut source = length_field(MODEL_STYLE_FIELD, &nested);
        source.extend_from_slice(&length_field(
            MODEL_STYLE_PRESET_FIELD,
            &reference_payload(9),
        ));
        source
    }
    fn unknown_group() -> Vec<u8> {
        vec![0x13, 0x08, 0x00, 0x14]
    }
    fn identified(name: &[u8], style_identifier: u64) -> Vec<u8> {
        let mut entry = length_field(IDENTIFIED_IDENTIFIER_FIELD, name);
        entry.extend_from_slice(&length_field(
            IDENTIFIED_STYLE_FIELD,
            &reference_payload(style_identifier),
        ));
        entry
    }
    fn style(parent: Option<u64>, props: AppearanceOverrides) -> Vec<u8> {
        let mut super_ = Vec::new();
        if let Some(parent) = parent {
            super_.extend_from_slice(&length_field(
                STYLE_PARENT_FIELD,
                &reference_payload(parent),
            ));
        }
        super_.extend_from_slice(&bool_field(STYLE_VARIATION_FIELD, parent.is_some()));
        super_.extend_from_slice(&length_field(STYLE_STYLESHEET_FIELD, &reference_payload(8)));
        let mut root = length_field(STYLE_SUPER_FIELD, &super_);
        let properties = canonical_properties(props);
        root.extend_from_slice(&length_field(STYLE_PROPERTIES_FIELD, &properties));
        root
    }

    #[test]
    fn model_style_edge_rewrite_preserves_unknown_bytes_and_reports_exact_limits() {
        let source = model(7, true);
        let (snapshot, _) = decode_table_model_with_report(&source, options(&source)).unwrap();
        assert_eq!(snapshot.style_identifier(), 7);
        let prepared =
            prepare_table_model_style_rewrite(&source, 7, 129, options(&source)).unwrap();
        let requirements = prepared.execution_requirements();
        let (candidate, report) = prepared.execute(requirements.exact_limits()).unwrap();
        assert_eq!(report.output_bytes(), candidate.len());
        assert_eq!(report.fields(), requirements.fields());
        assert_eq!(report.work_bytes(), requirements.work_bytes());
        assert!(
            candidate
                .windows(5)
                .any(|window| window == [0x98, 0x03, 0x81, 0x80, 0x00])
        );
        assert_eq!(
            decode_table_model(&candidate, options(&candidate))
                .unwrap()
                .style_identifier(),
            129
        );
        assert!(
            prepare_table_model_style_rewrite(
                &source,
                7,
                129,
                options(&source).with_max_output_bytes(candidate.len() - 1)
            )
            .is_err()
        );
        let prepared =
            prepare_table_model_style_rewrite(&source, 7, 129, options(&source)).unwrap();
        let requirements = prepared.execution_requirements();
        let mut low_allocations = requirements.exact_limits();
        low_allocations.max_allocations = requirements.allocations().saturating_sub(1);
        let error = prepared.execute(low_allocations).unwrap_err();
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Allocations { .. })
        ));
    }

    #[test]
    fn style_projection_inheritance_and_cycle_are_bounded() {
        let base_payload = style(
            None,
            AppearanceOverrides {
                body_horizontal: Some(false),
                ..Default::default()
            },
        );
        let child_payload = style(
            Some(9),
            AppearanceOverrides {
                row_banding: Some(true),
                ..Default::default()
            },
        );
        let base = decode_table_style(&base_payload, options(&base_payload)).unwrap();
        let child = decode_table_style(&child_payload, options(&child_payload)).unwrap();
        let nodes = [TableStyleNode::new(9, base), TableStyleNode::new(10, child)];
        let effective = resolve_table_style_appearance(&nodes, 10, options(&base_payload)).unwrap();
        assert!(effective.row_banding);
        assert!(!effective.body_horizontal);
        let cycle_payload = style(
            Some(10),
            AppearanceOverrides {
                row_banding: Some(true),
                ..Default::default()
            },
        );
        let cycle_style = decode_table_style(&cycle_payload, options(&cycle_payload)).unwrap();
        let cycle = [TableStyleNode::new(10, cycle_style)];
        assert!(resolve_table_style_appearance(&cycle, 10, options(&cycle_payload)).is_err());

        let duplicate_nodes = [
            TableStyleNode::new(10, child),
            TableStyleNode::new(10, base),
        ];
        assert!(
            resolve_table_style_appearance(&duplicate_nodes, 10, options(&base_payload)).is_err()
        );
    }

    #[test]
    fn canonical_variation_uses_strict_presence_and_limits() {
        let payload = canonical_table_style_variation(
            TableStyleVariationWrite {
                parent_identifier: 7,
                stylesheet_identifier: 8,
                overrides: AppearanceOverrides {
                    row_banding: Some(true),
                    row_sizing: Some(true),
                    body_horizontal: Some(true),
                    body_vertical: Some(false),
                    header_columns_horizontal: Some(true),
                    header_rows_vertical: Some(false),
                    footer_rows_vertical: Some(true),
                },
            },
            DecodeOptions::for_source(&[])
                .with_max_output_bytes(1024)
                .with_max_fields(64)
                .with_max_work_bytes(4096)
                .with_max_allocations(16),
        )
        .unwrap();
        assert!(!payload.bytes().is_empty());
        let decoded = decode_table_style_with_report(
            payload.bytes(),
            DecodeOptions::for_source(payload.bytes())
                .with_max_output_bytes(1024)
                .with_max_fields(64)
                .with_max_work_bytes(4096),
        )
        .unwrap();
        assert_eq!(payload.report().output_bytes(), decoded.1.input_bytes());
        assert_eq!(payload.report().fields(), decoded.1.fields());
        assert_eq!(payload.report().work_bytes(), decoded.1.work_bytes());
        assert_eq!(payload.report().max_depth(), decoded.1.max_depth());
        assert_eq!(decoded.0.parent_identifier(), Some(7));
        assert!(
            canonical_table_style_variation(
                TableStyleVariationWrite {
                    parent_identifier: 7,
                    stylesheet_identifier: 8,
                    overrides: AppearanceOverrides::default()
                },
                DecodeOptions::for_source(&[]).with_max_output_bytes(1)
            )
            .is_err()
        );
        assert!(matches!(
            canonical_table_style_variation(
                TableStyleVariationWrite {
                    parent_identifier: 7,
                    stylesheet_identifier: 8,
                    overrides: AppearanceOverrides {
                        row_banding: Some(true),
                        row_sizing: Some(true),
                        body_horizontal: Some(true),
                        body_vertical: Some(false),
                        header_columns_horizontal: Some(true),
                        header_rows_vertical: Some(false),
                        footer_rows_vertical: Some(true),
                    },
                },
                DecodeOptions::for_source(&[])
                    .with_max_output_bytes(1024)
                    .with_max_fields(64)
                    .with_max_work_bytes(4096)
                    .with_max_allocations(15),
            )
            .unwrap_err()
            .resource_limit(),
            Some(DecodeLimit::Allocations { .. })
        ));
    }

    #[test]
    fn stylesheet_append_preserves_source_and_checks_duplicate_style() {
        let source = vec![0x98, 0x03, 0x81, 0x80, 0x00];
        let append = StylesheetStyleAppend {
            style_identifier: 7,
            parent_identifier: None,
        };
        let (candidate, report) = append_stylesheet_style(
            &source,
            append,
            DecodeOptions::for_source(&source).with_max_output_bytes(1024),
        )
        .unwrap();
        assert!(candidate.starts_with(&source));
        assert_eq!(report.output_bytes(), candidate.len());
        assert_eq!(
            decode_stylesheet(&candidate, DecodeOptions::for_source(&candidate))
                .unwrap()
                .style_count(),
            1
        );
        assert!(
            append_stylesheet_style(
                &candidate,
                append,
                DecodeOptions::for_source(&candidate).with_max_output_bytes(1024)
            )
            .is_err()
        );
    }

    #[test]
    fn malformed_known_scalars_and_unbalanced_groups_fail_closed() {
        let mut source = model(7, false);
        source.extend_from_slice(&[0x18, 0x02]);
        assert!(decode_table_model(&source, options(&source)).is_err());
        assert!(
            decode_table_model(
                &[0x9b, 0x03, 0x08, 0x00],
                options(&[0x9b, 0x03, 0x08, 0x00])
            )
            .is_err()
        );
    }

    #[test]
    fn balanced_unknown_groups_before_and_after_known_edges_are_raw_preserved() {
        let mut source = unknown_group();
        source.extend_from_slice(&model(7, false));
        source.extend_from_slice(&unknown_group());
        let prepared = prepare_table_model_style_rewrite(
            &source,
            7,
            129,
            options(&source).with_max_work_bytes(4096),
        )
        .unwrap();
        let (candidate, report) = prepared
            .execute(prepared.execution_requirements().exact_limits())
            .unwrap();
        assert_eq!(report.output_bytes(), candidate.len());
        assert_eq!(candidate[..4], source[..4]);
        assert_eq!(
            &candidate[candidate.len() - 4..],
            &source[source.len() - 4..]
        );
        assert_eq!(
            decode_table_model(&candidate, options(&candidate))
                .unwrap()
                .style_identifier(),
            129
        );
    }

    #[test]
    fn balanced_unknown_groups_do_not_swallow_style_or_registry_fields() {
        let mut style_source = unknown_group();
        style_source.extend_from_slice(&style(
            None,
            AppearanceOverrides {
                body_horizontal: Some(false),
                ..Default::default()
            },
        ));
        style_source.extend_from_slice(&unknown_group());
        let style_snapshot = decode_table_style(&style_source, options(&style_source)).unwrap();
        assert_eq!(style_snapshot.raw(), style_source.as_slice());
        assert_eq!(style_snapshot.overrides().body_horizontal, Some(false));

        let stylesheet_unknown_group = || vec![0x33, 0x08, 0x00, 0x34];
        let mut stylesheet_source = stylesheet_unknown_group();
        stylesheet_source
            .extend_from_slice(&length_field(SHEET_STYLES_FIELD, &reference_payload(7)));
        stylesheet_source.extend_from_slice(&stylesheet_unknown_group());
        let append = StylesheetStyleAppend {
            style_identifier: 9,
            parent_identifier: None,
        };
        let prepared = prepare_stylesheet_append(
            &stylesheet_source,
            append,
            options(&stylesheet_source)
                .with_max_output_bytes(1024)
                .with_max_work_bytes(4096),
        )
        .unwrap();
        let (candidate, _) = prepared
            .execute(prepared.execution_requirements().exact_limits())
            .unwrap();
        assert_eq!(
            &candidate[..stylesheet_unknown_group().len()],
            stylesheet_unknown_group().as_slice()
        );
        assert_eq!(
            decode_stylesheet(&candidate, options(&candidate))
                .unwrap()
                .style_count(),
            2
        );
    }

    #[test]
    fn balanced_group_depth_is_a_typed_nesting_limit() {
        let mut source = Vec::new();
        source.extend(std::iter::repeat_n(0x13, 6));
        source.extend_from_slice(&[0x08, 0x00]);
        source.extend(std::iter::repeat_n(0x14, 6));
        source.extend_from_slice(&model(7, false));
        let error =
            decode_table_model(&source, options(&source).with_recursion_limit(3)).unwrap_err();
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Nesting { .. })
        ));
    }

    #[test]
    fn reference_known_fields_are_unique_local_and_canonical() {
        let mut duplicate_type = reference_payload(7);
        duplicate_type.extend_from_slice(&varint_field(2, 0));
        duplicate_type.extend_from_slice(&varint_field(2, 0));
        let source = length_field(MODEL_STYLE_FIELD, &duplicate_type);
        assert!(decode_table_model(&source, options(&source)).is_err());

        let mut external = reference_payload(7);
        external.extend_from_slice(&bool_field(3, true));
        let source = length_field(MODEL_STYLE_FIELD, &external);
        assert!(decode_table_model(&source, options(&source)).is_err());

        let mut unsupported = reference_payload(7);
        unsupported.extend_from_slice(&varint_field(2, 1));
        let source = length_field(MODEL_STYLE_FIELD, &unsupported);
        assert!(decode_table_model(&source, options(&source)).is_err());
    }

    #[test]
    fn stylesheet_registry_edges_are_global_and_parent_append_is_in_place() {
        let mut source = length_field(SHEET_STYLES_FIELD, &reference_payload(7));
        source.extend_from_slice(&length_field(SHEET_STYLES_FIELD, &reference_payload(9)));
        source.extend_from_slice(&length_field(SHEET_CHILDREN_FIELD, &children_entry(7, 9)));
        let append = StylesheetStyleAppend {
            style_identifier: 11,
            parent_identifier: Some(7),
        };
        let append_options = options(&source)
            .with_max_work_bytes(4096)
            .with_max_styles(3);
        let prepared = prepare_stylesheet_append(&source, append, append_options).unwrap();
        let requirements = prepared.execution_requirements();
        let (candidate, report) = prepared.execute(requirements.exact_limits()).unwrap();
        assert_eq!(report.output_bytes(), candidate.len());
        assert_eq!(report.fields(), requirements.fields());
        assert_eq!(report.work_bytes(), requirements.work_bytes());
        assert_eq!(report.allocations(), requirements.allocations());
        let scan = scan_stylesheet(&candidate, 256, 4096, 16, 3, None)
            .unwrap()
            .0;
        assert_eq!(scan.style_count, 3);
        assert_eq!(scan.child_entry_parent_len, 1);
        assert_eq!(scan.child_edges_len, 2);
        assert!(
            append_stylesheet_style(&source, append, append_options.with_max_styles(2),).is_err()
        );
        let prepared = prepare_stylesheet_append(&source, append, append_options).unwrap();
        let requirements = prepared.execution_requirements();
        let mut low_allocations = requirements.exact_limits();
        low_allocations.max_allocations = requirements.allocations().saturating_sub(1);
        let error = prepared.execute(low_allocations).unwrap_err();
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Allocations { .. })
        ));

        let mut bad_identified = length_field(SHEET_STYLES_FIELD, &reference_payload(7));
        bad_identified.extend_from_slice(&length_field(
            SHEET_IDENTIFIED_STYLES_FIELD,
            &identified(b"main", 8),
        ));
        assert!(decode_stylesheet(&bad_identified, options(&bad_identified)).is_err());

        let mut bad_child = length_field(SHEET_STYLES_FIELD, &reference_payload(7));
        bad_child.extend_from_slice(&length_field(SHEET_CHILDREN_FIELD, &children_entry(7, 8)));
        assert!(decode_stylesheet(&bad_child, options(&bad_child)).is_err());

        let mut versioned_payload = length_field(VERSIONED_STYLES_FIELD, &reference_payload(17));
        versioned_payload.extend_from_slice(&length_field(
            VERSIONED_IDENTIFIED_STYLES_FIELD,
            &identified(b"historical", 17),
        ));
        let mut versioned = length_field(SHEET_STYLES_FIELD, &reference_payload(7));
        versioned.extend_from_slice(&length_field(7, &versioned_payload));
        let prepared = prepare_stylesheet_append(
            &versioned,
            append,
            options(&versioned)
                .with_max_output_bytes(1024)
                .with_max_work_bytes(8192),
        )
        .unwrap();
        let (versioned_candidate, _) = prepared
            .execute(prepared.execution_requirements().exact_limits())
            .unwrap();
        assert!(
            versioned_candidate
                .windows(versioned_payload.len())
                .any(|window| window == versioned_payload.as_slice())
        );

        let reserved_append = StylesheetStyleAppend {
            style_identifier: 17,
            parent_identifier: Some(7),
        };
        assert!(
            prepare_stylesheet_append(
                &versioned,
                reserved_append,
                options(&versioned)
                    .with_max_output_bytes(1024)
                    .with_max_work_bytes(8192),
            )
            .is_err()
        );

        let mut duplicate_version = versioned.clone();
        duplicate_version.extend_from_slice(&length_field(7, &versioned_payload));
        assert!(decode_stylesheet(&duplicate_version, options(&duplicate_version)).is_err());
    }

    #[test]
    fn legacy_and_modern_header_overrides_must_agree() {
        let mut super_payload = length_field(STYLE_STYLESHEET_FIELD, &reference_payload(8));
        super_payload.extend_from_slice(&bool_field(STYLE_VARIATION_FIELD, true));
        let mut properties = bool_field(PROPERTIES_LEGACY_HEADER_COLUMN_FIELD, true);
        properties.extend_from_slice(&bool_field(PROPERTIES_HEADER_COLUMN_FIELD, true));
        let mut source = length_field(STYLE_SUPER_FIELD, &super_payload);
        source.extend_from_slice(&length_field(STYLE_PROPERTIES_FIELD, &properties));
        let snapshot = decode_table_style(&source, options(&source)).unwrap();
        assert_eq!(snapshot.overrides().header_columns_horizontal, Some(true));

        properties = bool_field(PROPERTIES_LEGACY_HEADER_COLUMN_FIELD, false);
        properties.extend_from_slice(&bool_field(PROPERTIES_HEADER_COLUMN_FIELD, true));
        let mut conflicting = length_field(STYLE_SUPER_FIELD, &super_payload);
        conflicting.extend_from_slice(&length_field(STYLE_PROPERTIES_FIELD, &properties));
        assert!(decode_table_style(&conflicting, options(&conflicting)).is_err());

        let mut duplicate_variation_super =
            length_field(STYLE_STYLESHEET_FIELD, &reference_payload(8));
        duplicate_variation_super.extend_from_slice(&bool_field(STYLE_VARIATION_FIELD, false));
        duplicate_variation_super.extend_from_slice(&bool_field(STYLE_VARIATION_FIELD, false));
        let duplicate_variation = length_field(STYLE_SUPER_FIELD, &duplicate_variation_super);
        assert!(decode_table_style(&duplicate_variation, options(&duplicate_variation)).is_err());

        let mut duplicate_count = length_field(STYLE_SUPER_FIELD, &super_payload);
        duplicate_count.extend_from_slice(&varint_field(STYLE_OVERRIDE_COUNT_FIELD, 7));
        duplicate_count.extend_from_slice(&varint_field(STYLE_OVERRIDE_COUNT_FIELD, 7));
        assert!(decode_table_style(&duplicate_count, options(&duplicate_count)).is_err());
    }
}
