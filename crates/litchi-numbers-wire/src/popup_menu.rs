//! Source-preserving BNC transitions for Numbers Pop-Up Menu cells.
//!
//! The general [`crate::BncCell`] adapter is convenient for older editor
//! paths, but it owns fields in a map and consequently emits the canonical
//! field sequence. A package transaction needs a smaller, stricter operation:
//! change only the format/control references and, when requested, the selected
//! string-table key while retaining every other source byte. This planner
//! borrows the source and performs validation and sizing before execution
//! allocates the candidate buffer.

use std::fmt;

const BNC_VERSION: u8 = 5;
const BNC_HEADER_LEN: usize = 12;
const BNC_PREFIX_LEN: usize = 8;

const CELL_TYPE_EMPTY: u8 = 0;
const CELL_TYPE_TEXT: u8 = 3;

const STRING_FLAG: u32 = 0x0000_0008;
const CONTROL_CELL_SPEC_FLAG: u32 = 0x0000_0400;
const CELL_FORMAT_KIND_FLAG: u32 = 0x0000_1000;
const TEXT_FORMAT_IDENTIFIER_FLAG: u32 = 0x0002_0000;
const FORMULA_FLAG: u32 = 0x0000_0200;

const EXPLICIT_TEXT_FORMAT: u16 = 0x0080;
const TEXT_CELL_FORMAT_KIND: u32 = 5;

// BNC v5 fields have no payload keys: flags select fixed-width values in this
// order. The opaque tail starts after the final selected known field.
const FIELD_LAYOUT: &[(u32, usize)] = &[
    (0x0000_0001, 16), // decimal
    (0x0000_0002, 8),  // number
    (0x0000_0004, 8),  // date
    (STRING_FLAG, 4),
    (0x0000_0010, 4), // rich text
    (0x0000_0020, 4), // style
    (0x0000_0040, 4), // text style
    (0x0000_0080, 4), // conditional style
    (0x0000_0100, 4), // conditional style rule
    (FORMULA_FLAG, 4),
    (CONTROL_CELL_SPEC_FLAG, 4),
    (0x0000_0800, 4), // formula error
    (CELL_FORMAT_KIND_FLAG, 4),
    (0x0000_2000, 4), // generic format identifier
    (0x0000_4000, 4), // currency format identifier
    (0x0000_8000, 4), // date-time format identifier
    (0x0001_0000, 4), // duration format identifier
    (TEXT_FORMAT_IDENTIFIER_FLAG, 4),
    (0x0004_0000, 4), // checkbox format identifier
    (0x0008_0000, 4), // comment
    (0x0010_0000, 4), // reserved known v5 field
];

const KNOWN_FLAGS: u32 = 0x001f_ffff;
const VALUE_FLAGS: u32 = 0x0000_020f;
const POPUP_METADATA_FLAGS: u32 =
    CONTROL_CELL_SPEC_FLAG | CELL_FORMAT_KIND_FLAG | TEXT_FORMAT_IDENTIFIER_FLAG;

/// The requested metadata/value state of one Pop-Up Menu BNC cell.
///
/// `set` requires both native references. `first_item_string_identifier` is
/// optional because a blank initial selection has no string-table value. A
/// `None` value means an existing string value is retained; this mirrors
/// Numbers' native behavior when applying a format to an already populated
/// text cell. [`Self::reset`] removes popup metadata while retaining the cell
/// value, including a selected string.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PopUpMenuBncState {
    data_format_identifier: Option<u32>,
    control_cell_spec_identifier: Option<u32>,
    first_item_string_identifier: Option<u32>,
}

impl PopUpMenuBncState {
    /// Construct a Pop-Up Menu metadata state.
    #[must_use]
    pub const fn set(
        data_format_identifier: u32,
        control_cell_spec_identifier: u32,
        first_item_string_identifier: Option<u32>,
    ) -> Self {
        Self {
            data_format_identifier: Some(data_format_identifier),
            control_cell_spec_identifier: Some(control_cell_spec_identifier),
            first_item_string_identifier,
        }
    }

    /// Construct the automatic/reset state.
    #[must_use]
    pub const fn reset() -> Self {
        Self {
            data_format_identifier: None,
            control_cell_spec_identifier: None,
            first_item_string_identifier: None,
        }
    }

    /// Native format-table entry key, if popup metadata is requested.
    #[must_use]
    pub const fn data_format_identifier(self) -> Option<u32> {
        self.data_format_identifier
    }

    /// Native control-cell-spec object key, if popup metadata is requested.
    #[must_use]
    pub const fn control_cell_spec_identifier(self) -> Option<u32> {
        self.control_cell_spec_identifier
    }

    /// Optional string-table key for the first menu item.
    #[must_use]
    pub const fn first_item_string_identifier(self) -> Option<u32> {
        self.first_item_string_identifier
    }

    #[must_use]
    const fn is_reset(self) -> bool {
        self.data_format_identifier.is_none() && self.control_cell_spec_identifier.is_none()
    }
}

/// Strict source-scan and planning limits.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PopUpMenuRewriteOptions {
    pub max_input_bytes: usize,
    pub max_output_bytes: usize,
    pub max_fields: usize,
    pub max_work_bytes: usize,
    pub recursion_limit: usize,
    pub max_references: usize,
}

impl Default for PopUpMenuRewriteOptions {
    fn default() -> Self {
        Self {
            max_input_bytes: 64 * 1024 * 1024,
            max_output_bytes: 64 * 1024 * 1024,
            max_fields: 100_000,
            max_work_bytes: 512 * 1024 * 1024,
            recursion_limit: 8,
            max_references: 100_000,
        }
    }
}

/// Exact source facts established without producing candidate bytes.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct PopUpMenuRewritePrepareReport {
    input_bytes: usize,
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: usize,
    references: usize,
}

impl PopUpMenuRewritePrepareReport {
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
    pub const fn max_depth(self) -> usize {
        self.max_depth
    }

    #[must_use]
    pub const fn references(self) -> usize {
        self.references
    }

    #[must_use]
    pub const fn allocations(self) -> usize {
        0
    }
}

/// Candidate execution requirements computed by preparation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PopUpMenuRewriteExecutionRequirements {
    output_bytes: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
    allocations: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: usize,
    references: usize,
}

impl PopUpMenuRewriteExecutionRequirements {
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
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
    pub const fn allocations(self) -> usize {
        self.allocations
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
    pub const fn max_depth(self) -> usize {
        self.max_depth
    }

    #[must_use]
    pub const fn references(self) -> usize {
        self.references
    }

    #[must_use]
    pub const fn exact_limits(self) -> PopUpMenuRewriteExecutionLimits {
        PopUpMenuRewriteExecutionLimits {
            max_output_bytes: self.output_bytes,
            max_retained_bytes: self.retained_bytes,
            max_scratch_bytes: self.scratch_bytes,
            max_allocations: self.allocations,
            max_fields: self.fields,
            max_work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            max_references: self.references,
        }
    }
}

/// Independent ceilings checked before candidate allocation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PopUpMenuRewriteExecutionLimits {
    pub max_output_bytes: usize,
    pub max_retained_bytes: usize,
    pub max_scratch_bytes: usize,
    pub max_allocations: usize,
    pub max_fields: usize,
    pub max_work_bytes: usize,
    pub max_depth: usize,
    pub max_references: usize,
}

/// Exact execution counters for a successful candidate.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PopUpMenuRewriteExecutionReport {
    retained_bytes: usize,
    scratch_bytes: usize,
    allocations: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: usize,
    references: usize,
}

impl PopUpMenuRewriteExecutionReport {
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }

    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
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
    pub const fn max_depth(self) -> usize {
        self.max_depth
    }

    #[must_use]
    pub const fn references(self) -> usize {
        self.references
    }
}

/// A validated candidate BNC payload.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PopUpMenuRewrite {
    bytes: Vec<u8>,
    report: PopUpMenuRewriteExecutionReport,
}

impl PopUpMenuRewrite {
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }

    #[must_use]
    pub const fn report(&self) -> PopUpMenuRewriteExecutionReport {
        self.report
    }
}

/// Errors returned by strict preparation and execution.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum PopUpMenuRewriteError {
    InvalidFormat(&'static str),
    InputBytes { observed: usize, maximum: usize },
    OutputBytes { observed: usize, maximum: usize },
    Fields { observed: usize, maximum: usize },
    Work { observed: usize, maximum: usize },
    Nesting { observed: usize, maximum: usize },
    References { observed: usize, maximum: usize },
    Allocations { observed: usize, maximum: usize },
    RetainedBytes { observed: usize, maximum: usize },
    ScratchBytes { observed: usize, maximum: usize },
    Allocation { requested: usize },
}

impl fmt::Display for PopUpMenuRewriteError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidFormat(message) => formatter.write_str(message),
            Self::InputBytes { observed, maximum } => {
                write!(formatter, "BNC input bytes {observed} exceed {maximum}")
            },
            Self::OutputBytes { observed, maximum } => {
                write!(formatter, "BNC output bytes {observed} exceed {maximum}")
            },
            Self::Fields { observed, maximum } => {
                write!(formatter, "BNC fields {observed} exceed {maximum}")
            },
            Self::Work { observed, maximum } => {
                write!(formatter, "BNC work bytes {observed} exceed {maximum}")
            },
            Self::Nesting { observed, maximum } => {
                write!(formatter, "BNC nesting {observed} exceed {maximum}")
            },
            Self::References { observed, maximum } => {
                write!(formatter, "BNC references {observed} exceed {maximum}")
            },
            Self::Allocations { observed, maximum } => {
                write!(formatter, "BNC allocations {observed} exceed {maximum}")
            },
            Self::RetainedBytes { observed, maximum } => {
                write!(formatter, "BNC retained bytes {observed} exceed {maximum}")
            },
            Self::ScratchBytes { observed, maximum } => {
                write!(formatter, "BNC scratch bytes {observed} exceed {maximum}")
            },
            Self::Allocation { requested } => {
                write!(formatter, "could not allocate {requested} BNC bytes")
            },
        }
    }
}

impl std::error::Error for PopUpMenuRewriteError {}

type Result<T> = std::result::Result<T, PopUpMenuRewriteError>;

#[derive(Clone, Copy)]
struct FieldRange {
    flag: u32,
    start: usize,
    end: usize,
}

struct ParsedCell<'a> {
    source: &'a [u8],
    fields: [Option<FieldRange>; FIELD_LAYOUT.len()],
    flags: u32,
    tail_start: usize,
    format_identifier: Option<u32>,
    control_identifier: Option<u32>,
    kind: Option<u32>,
    string_identifier: Option<u32>,
}

/// A strict, output-free Pop-Up Menu BNC planner.
pub struct PreparedPopUpMenuRewrite<'source> {
    source: &'source [u8],
    parsed: ParsedCell<'source>,
    target: PopUpMenuBncState,
    prepare_report: PopUpMenuRewritePrepareReport,
    requirements: PopUpMenuRewriteExecutionRequirements,
}

impl PreparedPopUpMenuRewrite<'_> {
    #[must_use]
    pub const fn prepare_report(&self) -> PopUpMenuRewritePrepareReport {
        self.prepare_report
    }

    #[must_use]
    pub const fn execution_requirements(&self) -> PopUpMenuRewriteExecutionRequirements {
        self.requirements
    }

    /// Execute after all independent ceilings are checked.
    pub fn execute(self, limits: PopUpMenuRewriteExecutionLimits) -> Result<PopUpMenuRewrite> {
        preflight(self.requirements, limits)?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(self.requirements.output_bytes)
            .map_err(|_| PopUpMenuRewriteError::Allocation {
                requested: self.requirements.output_bytes,
            })?;
        if bytes.capacity() != self.requirements.output_bytes {
            return Err(PopUpMenuRewriteError::Allocation {
                requested: self.requirements.output_bytes,
            });
        }
        emit(self.source, &self.parsed, self.target, &mut bytes)?;
        verify_candidate(self.source, &bytes, self.target)?;
        if bytes.len() != self.requirements.output_bytes {
            return Err(PopUpMenuRewriteError::InvalidFormat(
                "BNC candidate length changed during publication",
            ));
        }
        Ok(PopUpMenuRewrite {
            bytes,
            report: PopUpMenuRewriteExecutionReport {
                retained_bytes: self.requirements.retained_bytes,
                scratch_bytes: self.requirements.scratch_bytes,
                allocations: self.requirements.allocations,
                fields: self.requirements.fields,
                work_bytes: self.requirements.work_bytes,
                max_depth: self.requirements.max_depth,
                references: self.requirements.references,
            },
        })
    }
}

/// Prepare a source-preserving Pop-Up Menu metadata/value transition.
pub fn prepare_pop_up_menu_bnc<'source>(
    source: &'source [u8],
    target: PopUpMenuBncState,
    options: PopUpMenuRewriteOptions,
) -> Result<PreparedPopUpMenuRewrite<'source>> {
    if source.len() > options.max_input_bytes {
        return Err(PopUpMenuRewriteError::InputBytes {
            observed: source.len(),
            maximum: options.max_input_bytes,
        });
    }
    let parsed = parse(source, options)?;
    validate_transition(&parsed, target)?;
    let output_bytes = output_len(&parsed, target)?;
    let source_fields = parsed.fields.iter().flatten().count();
    let source_references = usize::from(parsed.format_identifier.is_some())
        + usize::from(parsed.control_identifier.is_some())
        + usize::from(parsed.string_identifier.is_some());
    let prepare_report = PopUpMenuRewritePrepareReport {
        input_bytes: source.len(),
        output_bytes,
        fields: source_fields,
        work_bytes: source.len(),
        max_depth: 0,
        references: source_references,
    };
    let (candidate_fields, candidate_references) = candidate_facts(&parsed, target);
    let fields = source_fields
        .checked_mul(2)
        .and_then(|value| value.checked_add(candidate_fields))
        .ok_or(PopUpMenuRewriteError::Fields {
            observed: usize::MAX,
            maximum: options.max_fields,
        })?;
    let references = source_references
        .checked_mul(2)
        .and_then(|value| value.checked_add(candidate_references))
        .ok_or(PopUpMenuRewriteError::References {
            observed: usize::MAX,
            maximum: options.max_references,
        })?;
    let tail_bytes = source.len().saturating_sub(parsed.tail_start);
    let work_bytes = source
        .len()
        .checked_mul(2)
        .and_then(|value| value.checked_add(output_bytes.checked_mul(2)?))
        .and_then(|value| value.checked_add(tail_bytes.checked_mul(2)?))
        .ok_or(PopUpMenuRewriteError::Work {
            observed: usize::MAX,
            maximum: options.max_work_bytes,
        })?;
    let requirements = PopUpMenuRewriteExecutionRequirements {
        output_bytes,
        retained_bytes: output_bytes,
        scratch_bytes: 0,
        allocations: 1,
        fields,
        work_bytes,
        max_depth: 0,
        references,
    };
    check_report(prepare_report, requirements, options)?;
    Ok(PreparedPopUpMenuRewrite {
        source,
        parsed,
        target,
        prepare_report,
        requirements,
    })
}

/// One-shot convenience wrapper using exact execution requirements.
pub fn rewrite_pop_up_menu_bnc(
    source: &[u8],
    target: PopUpMenuBncState,
    options: PopUpMenuRewriteOptions,
) -> Result<PopUpMenuRewrite> {
    let prepared = prepare_pop_up_menu_bnc(source, target, options)?;
    let limits = prepared.execution_requirements().exact_limits();
    prepared.execute(limits)
}

// Short aliases make the low-level seam easy to discover while retaining a
// descriptive canonical name for package code and boundary audits.
pub use prepare_pop_up_menu_bnc as prepare_bnc_pop_up_menu_transition;
pub use rewrite_pop_up_menu_bnc as rewrite_bnc_pop_up_menu_transition;

fn parse<'a>(source: &'a [u8], options: PopUpMenuRewriteOptions) -> Result<ParsedCell<'a>> {
    if source.len() < BNC_HEADER_LEN {
        return Err(PopUpMenuRewriteError::InvalidFormat(
            "truncated Numbers BNC cell header",
        ));
    }
    if source[0] != BNC_VERSION {
        return Err(PopUpMenuRewriteError::InvalidFormat(
            "Numbers BNC cell is not writable version 5",
        ));
    }
    let flags = u32::from_le_bytes(source[8..12].try_into().expect("validated header"));
    if flags & !KNOWN_FLAGS != 0 {
        return Err(PopUpMenuRewriteError::InvalidFormat(
            "Numbers BNC cell contains unknown field flags",
        ));
    }
    let mut fields = [None; FIELD_LAYOUT.len()];
    let mut cursor = BNC_HEADER_LEN;
    for (index, &(flag, size)) in FIELD_LAYOUT.iter().enumerate() {
        if flags & flag == 0 {
            continue;
        }
        let end = cursor
            .checked_add(size)
            .ok_or(PopUpMenuRewriteError::InvalidFormat(
                "Numbers BNC field offset overflow",
            ))?;
        if end > source.len() {
            return Err(PopUpMenuRewriteError::InvalidFormat(
                "truncated Numbers BNC field",
            ));
        }
        fields[index] = Some(FieldRange {
            flag,
            start: cursor,
            end,
        });
        cursor = end;
    }
    let parsed = ParsedCell {
        source,
        fields,
        flags,
        tail_start: cursor,
        format_identifier: read_u32(&fields, source, TEXT_FORMAT_IDENTIFIER_FLAG),
        control_identifier: read_u32(&fields, source, CONTROL_CELL_SPEC_FLAG),
        kind: read_u32(&fields, source, CELL_FORMAT_KIND_FLAG),
        string_identifier: read_u32(&fields, source, STRING_FLAG),
    };
    let field_count = parsed.fields.iter().flatten().count();
    if field_count > options.max_fields {
        return Err(PopUpMenuRewriteError::Fields {
            observed: field_count,
            maximum: options.max_fields,
        });
    }
    Ok(parsed)
}

fn read_u32(
    fields: &[Option<FieldRange>; FIELD_LAYOUT.len()],
    source: &[u8],
    flag: u32,
) -> Option<u32> {
    let field = fields.iter().flatten().find(|field| field.flag == flag)?;
    let bytes = source.get(field.start..field.end)?;
    Some(u32::from_le_bytes(bytes.try_into().ok()?))
}

fn validate_transition(parsed: &ParsedCell<'_>, target: PopUpMenuBncState) -> Result<()> {
    if target.data_format_identifier.is_some() != target.control_cell_spec_identifier.is_some() {
        return Err(PopUpMenuRewriteError::InvalidFormat(
            "Pop-Up Menu format and control references must be present together",
        ));
    }
    if target
        .data_format_identifier
        .is_some_and(|identifier| identifier == 0)
        || target
            .control_cell_spec_identifier
            .is_some_and(|identifier| identifier == 0)
        || target
            .first_item_string_identifier
            .is_some_and(|identifier| identifier == 0)
    {
        return Err(PopUpMenuRewriteError::InvalidFormat(
            "Pop-Up Menu references must be nonzero",
        ));
    }

    let value_flags = parsed.flags & VALUE_FLAGS;
    if value_flags & !STRING_FLAG != 0 {
        return Err(PopUpMenuRewriteError::InvalidFormat(
            "Pop-Up Menu transition would discard a non-text BNC value",
        ));
    }
    if value_flags & STRING_FLAG != 0 && parsed.string_identifier.is_none() {
        return Err(PopUpMenuRewriteError::InvalidFormat(
            "BNC string value has no string-table identifier",
        ));
    }
    if value_flags & STRING_FLAG != 0 && parsed.source[1] != CELL_TYPE_TEXT {
        return Err(PopUpMenuRewriteError::InvalidFormat(
            "BNC string value has a non-text cell type",
        ));
    }
    if target.is_reset() {
        if let Some(kind) = parsed.kind
            && kind != TEXT_CELL_FORMAT_KIND
        {
            return Err(PopUpMenuRewriteError::InvalidFormat(
                "Pop-Up Menu reset targeted a non-text format",
            ));
        }
    } else {
        if parsed.source[1] != CELL_TYPE_EMPTY && parsed.source[1] != CELL_TYPE_TEXT {
            return Err(PopUpMenuRewriteError::InvalidFormat(
                "Pop-Up Menu transition targeted a non-empty/non-text BNC cell",
            ));
        }
        if let Some(kind) = parsed.kind
            && kind != TEXT_CELL_FORMAT_KIND
        {
            return Err(PopUpMenuRewriteError::InvalidFormat(
                "Pop-Up Menu transition targeted a non-text format",
            ));
        }
    }
    Ok(())
}

fn output_len(parsed: &ParsedCell<'_>, target: PopUpMenuBncState) -> Result<usize> {
    let mut output = parsed.source.len();
    let old_flags = parsed.flags;
    let new_flags = target_flags(parsed, target);
    for &(flag, width) in FIELD_LAYOUT {
        let old_present = old_flags & flag != 0;
        let new_present = new_flags & flag != 0;
        if old_present == new_present {
            continue;
        }
        output = if new_present {
            output.checked_add(width)
        } else {
            output.checked_sub(width)
        }
        .ok_or(PopUpMenuRewriteError::InvalidFormat(
            "BNC output length overflow",
        ))?;
    }
    Ok(output)
}

fn target_flags(parsed: &ParsedCell<'_>, target: PopUpMenuBncState) -> u32 {
    let old_flags = parsed.flags;
    let new_string = if target.is_reset() {
        parsed.string_identifier.is_some()
    } else {
        target.first_item_string_identifier.is_some() || parsed.string_identifier.is_some()
    };
    if target.is_reset() {
        old_flags & !POPUP_METADATA_FLAGS
    } else if new_string {
        old_flags | POPUP_METADATA_FLAGS | STRING_FLAG
    } else {
        old_flags | POPUP_METADATA_FLAGS
    }
}

fn candidate_facts(parsed: &ParsedCell<'_>, target: PopUpMenuBncState) -> (usize, usize) {
    let flags = target_flags(parsed, target);
    let fields = FIELD_LAYOUT
        .iter()
        .filter(|(flag, _)| flags & flag != 0)
        .count();
    let references = usize::from(target.data_format_identifier.is_some())
        + usize::from(target.control_cell_spec_identifier.is_some())
        + usize::from(if target.is_reset() {
            parsed.string_identifier.is_some()
        } else {
            target.first_item_string_identifier.is_some() || parsed.string_identifier.is_some()
        });
    (fields, references)
}

fn check_report(
    prepare_report: PopUpMenuRewritePrepareReport,
    requirements: PopUpMenuRewriteExecutionRequirements,
    options: PopUpMenuRewriteOptions,
) -> Result<()> {
    if prepare_report.input_bytes > options.max_input_bytes {
        return Err(PopUpMenuRewriteError::InputBytes {
            observed: prepare_report.input_bytes,
            maximum: options.max_input_bytes,
        });
    }
    if requirements.output_bytes > options.max_output_bytes {
        return Err(PopUpMenuRewriteError::OutputBytes {
            observed: requirements.output_bytes,
            maximum: options.max_output_bytes,
        });
    }
    if requirements.fields > options.max_fields {
        return Err(PopUpMenuRewriteError::Fields {
            observed: requirements.fields,
            maximum: options.max_fields,
        });
    }
    if requirements.work_bytes > options.max_work_bytes {
        return Err(PopUpMenuRewriteError::Work {
            observed: requirements.work_bytes,
            maximum: options.max_work_bytes,
        });
    }
    if requirements.max_depth > options.recursion_limit {
        return Err(PopUpMenuRewriteError::Nesting {
            observed: requirements.max_depth,
            maximum: options.recursion_limit,
        });
    }
    if requirements.references > options.max_references {
        return Err(PopUpMenuRewriteError::References {
            observed: requirements.references,
            maximum: options.max_references,
        });
    }
    Ok(())
}

fn preflight(
    requirements: PopUpMenuRewriteExecutionRequirements,
    limits: PopUpMenuRewriteExecutionLimits,
) -> Result<()> {
    if requirements.output_bytes > limits.max_output_bytes {
        return Err(PopUpMenuRewriteError::OutputBytes {
            observed: requirements.output_bytes,
            maximum: limits.max_output_bytes,
        });
    }
    if requirements.retained_bytes > limits.max_retained_bytes {
        return Err(PopUpMenuRewriteError::RetainedBytes {
            observed: requirements.retained_bytes,
            maximum: limits.max_retained_bytes,
        });
    }
    if requirements.scratch_bytes > limits.max_scratch_bytes {
        return Err(PopUpMenuRewriteError::ScratchBytes {
            observed: requirements.scratch_bytes,
            maximum: limits.max_scratch_bytes,
        });
    }
    if requirements.allocations > limits.max_allocations {
        return Err(PopUpMenuRewriteError::Allocations {
            observed: requirements.allocations,
            maximum: limits.max_allocations,
        });
    }
    if requirements.fields > limits.max_fields {
        return Err(PopUpMenuRewriteError::Fields {
            observed: requirements.fields,
            maximum: limits.max_fields,
        });
    }
    if requirements.work_bytes > limits.max_work_bytes {
        return Err(PopUpMenuRewriteError::Work {
            observed: requirements.work_bytes,
            maximum: limits.max_work_bytes,
        });
    }
    if requirements.max_depth > limits.max_depth {
        return Err(PopUpMenuRewriteError::Nesting {
            observed: requirements.max_depth,
            maximum: limits.max_depth,
        });
    }
    if requirements.references > limits.max_references {
        return Err(PopUpMenuRewriteError::References {
            observed: requirements.references,
            maximum: limits.max_references,
        });
    }
    Ok(())
}

fn emit(
    source: &[u8],
    parsed: &ParsedCell<'_>,
    target: PopUpMenuBncState,
    output: &mut Vec<u8>,
) -> Result<()> {
    let mut prefix = [0u8; BNC_PREFIX_LEN];
    prefix.copy_from_slice(&source[..BNC_PREFIX_LEN]);
    let mut flags = parsed.flags;
    if target.is_reset() {
        prefix[6..8].copy_from_slice(&0u16.to_le_bytes());
        flags &= !POPUP_METADATA_FLAGS;
    } else {
        prefix[6..8].copy_from_slice(&EXPLICIT_TEXT_FORMAT.to_le_bytes());
        flags |= POPUP_METADATA_FLAGS;
    }
    let old_string = parsed.string_identifier;
    let desired_string = if target.is_reset() {
        old_string
    } else {
        target.first_item_string_identifier.or(old_string)
    };
    if desired_string.is_some() {
        flags |= STRING_FLAG;
    } else {
        flags &= !STRING_FLAG;
    }
    if desired_string.is_some() {
        prefix[1] = CELL_TYPE_TEXT;
    }
    output.extend_from_slice(&prefix);
    output.extend_from_slice(&flags.to_le_bytes());
    for &(flag, size) in FIELD_LAYOUT {
        if flags & flag == 0 {
            continue;
        }
        if let Some(identifier) = desired_identifier(flag, target, desired_string) {
            output.extend_from_slice(&identifier.to_le_bytes());
        } else if flag == CELL_FORMAT_KIND_FLAG {
            output.extend_from_slice(&TEXT_CELL_FORMAT_KIND.to_le_bytes());
        } else if let Some(field) = parsed
            .fields
            .iter()
            .flatten()
            .find(|field| field.flag == flag)
        {
            output.extend_from_slice(source.get(field.start..field.end).ok_or(
                PopUpMenuRewriteError::InvalidFormat("BNC retained field range is invalid"),
            )?);
        } else {
            return Err(PopUpMenuRewriteError::InvalidFormat(
                "BNC output field is missing from source",
            ));
        }
        if size != 4 && (flag == STRING_FLAG || flag == CONTROL_CELL_SPEC_FLAG) {
            return Err(PopUpMenuRewriteError::InvalidFormat(
                "BNC popup reference width is invalid",
            ));
        }
    }
    output.extend_from_slice(source.get(parsed.tail_start..).ok_or(
        PopUpMenuRewriteError::InvalidFormat("BNC opaque tail range is invalid"),
    )?);
    Ok(())
}

fn desired_identifier(
    flag: u32,
    target: PopUpMenuBncState,
    desired_string: Option<u32>,
) -> Option<u32> {
    match flag {
        TEXT_FORMAT_IDENTIFIER_FLAG => target.data_format_identifier,
        CONTROL_CELL_SPEC_FLAG => target.control_cell_spec_identifier,
        STRING_FLAG => desired_string,
        _ => None,
    }
}

fn verify_candidate(source: &[u8], candidate: &[u8], target: PopUpMenuBncState) -> Result<()> {
    let options = PopUpMenuRewriteOptions {
        max_input_bytes: candidate.len(),
        max_output_bytes: candidate.len(),
        max_fields: usize::MAX,
        max_work_bytes: usize::MAX,
        recursion_limit: usize::MAX,
        max_references: usize::MAX,
    };
    let parsed = parse(candidate, options)?;
    validate_transition(&parsed, target)?;
    if parsed.format_identifier != target.data_format_identifier
        || parsed.control_identifier != target.control_cell_spec_identifier
    {
        return Err(PopUpMenuRewriteError::InvalidFormat(
            "BNC candidate popup references differ from requested state",
        ));
    }
    if !target.is_reset()
        && target.first_item_string_identifier.is_some()
        && parsed.string_identifier != target.first_item_string_identifier
    {
        return Err(PopUpMenuRewriteError::InvalidFormat(
            "BNC candidate string value differs from requested state",
        ));
    }
    let source_parsed = parse(
        source,
        PopUpMenuRewriteOptions {
            max_input_bytes: source.len(),
            max_output_bytes: usize::MAX,
            max_fields: usize::MAX,
            max_work_bytes: usize::MAX,
            recursion_limit: usize::MAX,
            max_references: usize::MAX,
        },
    )?;
    let source_tail =
        source
            .get(source_parsed.tail_start..)
            .ok_or(PopUpMenuRewriteError::InvalidFormat(
                "BNC source opaque tail range is invalid",
            ))?;
    let candidate_tail =
        candidate
            .get(parsed.tail_start..)
            .ok_or(PopUpMenuRewriteError::InvalidFormat(
                "BNC candidate opaque tail range is invalid",
            ))?;
    if source_tail != candidate_tail {
        return Err(PopUpMenuRewriteError::InvalidFormat(
            "BNC candidate changed opaque tail bytes",
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn source_with_tail() -> Vec<u8> {
        let mut source = vec![BNC_VERSION, CELL_TYPE_EMPTY, 0xa5, 0x5a, 0x13, 0x37, 0, 0];
        let flags: u32 = 0x0000_0020 | 0x0000_0040;
        source.extend_from_slice(&flags.to_le_bytes());
        source.extend_from_slice(&17u32.to_le_bytes());
        source.extend_from_slice(&29u32.to_le_bytes());
        source.extend_from_slice(b"opaque-tail-with-unknown-order");
        source
    }

    fn options() -> PopUpMenuRewriteOptions {
        PopUpMenuRewriteOptions {
            max_input_bytes: usize::MAX,
            max_output_bytes: usize::MAX,
            max_fields: usize::MAX,
            max_work_bytes: usize::MAX,
            recursion_limit: usize::MAX,
            max_references: usize::MAX,
        }
    }

    #[test]
    fn prepared_set_preserves_unrelated_prefix_fields_and_tail() {
        let source = source_with_tail();
        let target = PopUpMenuBncState::set(31, 47, Some(61));
        let prepared = prepare_pop_up_menu_bnc(&source, target, options()).unwrap();
        assert_eq!(prepared.prepare_report().input_bytes(), source.len());
        assert_eq!(
            prepared.prepare_report().output_bytes(),
            prepared.execution_requirements().output_bytes()
        );
        let limits = prepared.execution_requirements().exact_limits();
        let output = prepared.execute(limits).unwrap().into_bytes();
        assert_eq!(&output[2..6], &source[2..6]);
        assert!(output.ends_with(b"opaque-tail-with-unknown-order"));
        let parsed = parse(&output, options()).unwrap();
        assert_eq!(parsed.format_identifier, Some(31));
        assert_eq!(parsed.control_identifier, Some(47));
        assert_eq!(parsed.kind, Some(TEXT_CELL_FORMAT_KIND));
        assert_eq!(parsed.string_identifier, Some(61));
    }

    #[test]
    fn prepared_reset_removes_popup_metadata_but_retains_value_and_tail() {
        let source = rewrite_pop_up_menu_bnc(
            &source_with_tail(),
            PopUpMenuBncState::set(31, 47, Some(61)),
            options(),
        )
        .unwrap()
        .into_bytes();
        let output = rewrite_pop_up_menu_bnc(&source, PopUpMenuBncState::reset(), options())
            .unwrap()
            .into_bytes();
        let parsed = parse(&output, options()).unwrap();
        assert_eq!(parsed.format_identifier, None);
        assert_eq!(parsed.control_identifier, None);
        assert_eq!(parsed.kind, None);
        assert_eq!(parsed.string_identifier, Some(61));
        assert_eq!(output[1], CELL_TYPE_TEXT);
        assert!(output.ends_with(b"opaque-tail-with-unknown-order"));
    }

    #[test]
    fn exact_and_minus_one_output_limits_are_checked_before_execute_allocation() {
        let source = source_with_tail();
        let target = PopUpMenuBncState::set(31, 47, Some(61));
        let prepared = prepare_pop_up_menu_bnc(&source, target, options()).unwrap();
        let required = prepared.execution_requirements();
        assert!(prepared.execute(required.exact_limits()).is_ok());

        let prepared = prepare_pop_up_menu_bnc(&source, target, options()).unwrap();
        let mut too_small = prepared.execution_requirements().exact_limits();
        too_small.max_output_bytes -= 1;
        assert!(matches!(
            prepared.execute(too_small),
            Err(PopUpMenuRewriteError::OutputBytes { .. })
        ));
    }

    #[test]
    fn exact_and_minus_one_work_and_field_limits_are_typed() {
        let source = source_with_tail();
        let target = PopUpMenuBncState::set(31, 47, Some(61));
        let prepared = prepare_pop_up_menu_bnc(&source, target, options()).unwrap();
        let requirements = prepared.execution_requirements();
        let limits = requirements.exact_limits();
        let result = prepared.execute(limits).unwrap();
        let report = result.report();
        assert_eq!(report.fields(), requirements.fields());
        assert_eq!(report.work_bytes(), requirements.work_bytes());
        assert_eq!(report.references(), requirements.references());
        assert_eq!(report.retained_bytes(), requirements.retained_bytes());
        assert_eq!(report.scratch_bytes(), requirements.scratch_bytes());
        assert_eq!(report.allocations(), requirements.allocations());

        let prepared = prepare_pop_up_menu_bnc(&source, target, options()).unwrap();
        let mut fields = prepared.execution_requirements().exact_limits();
        fields.max_fields -= 1;
        assert!(matches!(
            prepared.execute(fields),
            Err(PopUpMenuRewriteError::Fields { .. })
        ));

        let prepared = prepare_pop_up_menu_bnc(&source, target, options()).unwrap();
        let mut work = prepared.execution_requirements().exact_limits();
        work.max_work_bytes -= 1;
        assert!(matches!(
            prepared.execute(work),
            Err(PopUpMenuRewriteError::Work { .. })
        ));

        let prepared = prepare_pop_up_menu_bnc(&source, target, options()).unwrap();
        let mut references = prepared.execution_requirements().exact_limits();
        references.max_references -= 1;
        assert!(matches!(
            prepared.execute(references),
            Err(PopUpMenuRewriteError::References { .. })
        ));

        let prepared = prepare_pop_up_menu_bnc(&source, target, options()).unwrap();
        let mut allocations = prepared.execution_requirements().exact_limits();
        allocations.max_allocations -= 1;
        assert!(matches!(
            prepared.execute(allocations),
            Err(PopUpMenuRewriteError::Allocations { .. })
        ));

        let prepared = prepare_pop_up_menu_bnc(&source, target, options()).unwrap();
        let mut retained = prepared.execution_requirements().exact_limits();
        retained.max_retained_bytes -= 1;
        assert!(matches!(
            prepared.execute(retained),
            Err(PopUpMenuRewriteError::RetainedBytes { .. })
        ));

        let prepared = prepare_pop_up_menu_bnc(&source, target, options()).unwrap();
        let scratch = prepared.execution_requirements().exact_limits();
        assert_eq!(scratch.max_scratch_bytes, 0);
        assert!(prepared.execute(scratch).is_ok());
    }

    #[test]
    fn malformed_and_unsafe_transitions_fail_closed() {
        let mut unknown = source_with_tail();
        let flags =
            (u32::from_le_bytes(unknown[8..12].try_into().unwrap()) | 0x8000_0000).to_le_bytes();
        unknown[8..12].copy_from_slice(&flags);
        assert!(matches!(
            prepare_pop_up_menu_bnc(&unknown, PopUpMenuBncState::set(31, 47, None), options()),
            Err(PopUpMenuRewriteError::InvalidFormat(_))
        ));
        let mut number = source_with_tail();
        number[1] = 2;
        assert!(matches!(
            prepare_pop_up_menu_bnc(&number, PopUpMenuBncState::set(31, 47, None), options()),
            Err(PopUpMenuRewriteError::InvalidFormat(_))
        ));
        assert!(matches!(
            prepare_pop_up_menu_bnc(
                &source_with_tail(),
                PopUpMenuBncState::set(0, 47, None),
                options()
            ),
            Err(PopUpMenuRewriteError::InvalidFormat(_))
        ));
    }
}
