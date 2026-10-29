//! The chart-font graph engine (record 0375a): a chart's uniform font
//! identity, face traits and point size, stored in the paragraph styles its
//! `paragraph_styles` list (`TSCH.ChartArchive`, field 20) names.
//!
//! [`set`] (with a reset [`Change`]) is the host's `set_chart_font*` /
//! `reset_chart_font*` ported operation for operation over a
//! [`Workspace`]: every unique paragraph style of the chart is replaced by a
//! fresh variation (of its parent when it is a font-only variation, else of
//! itself) holding the font overrides that differ from the inherited values,
//! or by its parent when none is left; the chart's references are repointed,
//! the variations registered (stylesheet, UUID, external reference), and the
//! replaced font-only variations no longer used are removed. The variation
//! payloads are `litchi_iwa_text_wire::paragraph_style::font_variation`.

use std::collections::{BTreeMap, HashSet};
use std::fmt;
use std::mem::size_of;

use litchi_iwa_common::wire::{self, WireField};
use litchi_iwa_common::{LimitKind as WireLimitKind, WireLimits};
use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_iwa_protos::chart_style_codec::{self as codec, Limits as CodecLimits};
use litchi_iwa_text::character::{TextPointSize, TextStyle};
use litchi_iwa_text::font::Font;
use litchi_iwa_text_wire::paragraph_style::{
    self as text_style, FontOverrides, Property, PropertyKind, Styles,
};

use crate::graph::{self, Workspace};
use crate::shape;
use thiserror::Error as ThisError;

const CHART_MESSAGE_TYPE: u32 = 5_021;
const PARAGRAPH_STYLE_MESSAGE_TYPE: u32 = 2_022;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const CHART_EXTENSION_FIELD: u32 = 10_000;
const CHART_PARAGRAPH_STYLES_FIELD: u32 = 20;
const STYLE_SUPER_FIELD: u32 = 1;
const STYLE_PARENT_FIELD: u32 = 3;
const STYLE_STYLESHEET_FIELD: u32 = 5;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;

/// A chart's font identity and face traits.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ChartFont {
    /// The font.
    pub font: Font,
    /// Bold.
    pub bold: bool,
    /// Italic.
    pub italic: bool,
}

/// What a chart-font write changes.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum Change {
    /// Set the font identity and face traits.
    Font(ChartFont),
    /// Set the point size.
    Size(TextPointSize),
    /// Reset the font identity and face traits.
    ResetFont,
    /// Reset the point size.
    ResetSize,
}

/// The finite resources admitted by one chart-font operation.
///
/// The format owner derives this profile from its residual package budget;
/// the engine never restores a process-wide default after a caller supplies a
/// narrower profile.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    max_input_bytes: usize,
    max_fields: usize,
    max_output_bytes: usize,
    max_nesting: u32,
    max_work_bytes: usize,
    max_text_bytes: usize,
    max_allocations: usize,
    max_retained_bytes: usize,
    max_scratch_bytes: usize,
}

impl Limits {
    /// Build an explicit finite chart-font policy.
    #[must_use]
    pub const fn new(
        max_input_bytes: usize,
        max_fields: usize,
        max_output_bytes: usize,
        max_nesting: u32,
        max_work_bytes: usize,
        max_text_bytes: usize,
        max_allocations: usize,
        max_retained_bytes: usize,
        max_scratch_bytes: usize,
    ) -> Self {
        Self {
            max_input_bytes,
            max_fields,
            max_output_bytes,
            max_nesting,
            max_work_bytes,
            max_text_bytes,
            max_allocations,
            max_retained_bytes,
            max_scratch_bytes,
        }
    }

    /// Derive a conservative finite profile from one chart payload.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        Self::new(
            bytes.saturating_mul(256).max(1),
            bytes.saturating_mul(128).max(1),
            bytes.saturating_mul(128).max(1),
            64,
            bytes.saturating_mul(2048).max(1),
            bytes.saturating_mul(8).max(1),
            bytes.saturating_mul(1024).max(8),
            bytes.saturating_mul(256).max(1),
            bytes.saturating_mul(128).max(1),
        )
    }

    /// Maximum complete source payload bytes admitted.
    #[must_use]
    pub const fn max_input_bytes(self) -> usize {
        self.max_input_bytes
    }

    /// Maximum aggregate wire fields admitted.
    #[must_use]
    pub const fn max_fields(self) -> usize {
        self.max_fields
    }

    /// Maximum candidate payload bytes admitted.
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }

    /// Maximum paragraph-style inheritance depth.
    #[must_use]
    pub const fn max_nesting(self) -> u32 {
        self.max_nesting
    }

    /// Maximum aggregate graph and wire work.
    #[must_use]
    pub const fn max_work_bytes(self) -> usize {
        self.max_work_bytes
    }

    /// Maximum retained font-name bytes.
    #[must_use]
    pub const fn max_text_bytes(self) -> usize {
        self.max_text_bytes
    }

    /// Maximum admitted collection allocations.
    #[must_use]
    pub const fn max_allocations(self) -> usize {
        self.max_allocations
    }

    /// Maximum source and candidate bytes retained together.
    #[must_use]
    pub const fn max_retained_bytes(self) -> usize {
        self.max_retained_bytes
    }

    /// Maximum scratch bytes reserved while preparing a candidate.
    #[must_use]
    pub const fn max_scratch_bytes(self) -> usize {
        self.max_scratch_bytes
    }
}

/// Resource dimension reported by the chart-font engine.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete chart payload bytes inspected.
    InputBytes,
    /// Strict wire fields visited.
    Fields,
    /// Candidate chart/style payload bytes emitted.
    OutputBytes,
    /// Paragraph-style inheritance depth.
    Nesting,
    /// Aggregate graph and wire work.
    Work,
    /// Font-name bytes retained.
    TextBytes,
    /// Collection allocations.
    Allocations,
    /// Source and candidate bytes retained.
    RetainedBytes,
    /// Candidate scratch bytes.
    ScratchBytes,
    /// Physical archive resource from the graph layer.
    Archive(litchi_iwa_core::LimitKind),
    /// Wire resource from a graph adapter.
    Wire(WireLimitKind),
    /// Generated chart-style codec resource.
    Codec(codec::LimitKind),
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InputBytes => formatter.write_str("input bytes"),
            Self::Fields => formatter.write_str("wire fields"),
            Self::OutputBytes => formatter.write_str("output bytes"),
            Self::Nesting => formatter.write_str("wire nesting"),
            Self::Work => formatter.write_str("work"),
            Self::TextBytes => formatter.write_str("text bytes"),
            Self::Allocations => formatter.write_str("allocations"),
            Self::RetainedBytes => formatter.write_str("retained bytes"),
            Self::ScratchBytes => formatter.write_str("scratch bytes"),
            Self::Archive(kind) => write!(formatter, "archive {kind}"),
            Self::Wire(kind) => write!(formatter, "wire {kind}"),
            Self::Codec(kind) => write!(formatter, "codec {kind:?}"),
        }
    }
}

/// A typed chart-font engine failure.
#[derive(Debug, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// The selected chart graph or payload is malformed.
    #[error("invalid chart-font source: {0}")]
    InvalidSource(&'static str),
    /// A finite operation resource was exhausted.
    #[error("chart font {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource dimension.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded collection allocation failed.
    #[error("could not allocate {amount} chart-font units")]
    Allocation {
        /// Requested bytes or collection items.
        amount: usize,
    },
    /// A text-wire operation refused a well-formed request.
    #[error("chart-font request refused: {0:?}")]
    Refused(litchi_iwa_text_wire::paragraph_style::Refusal),
    /// A graph operation failed after preserving its typed resource details.
    #[error("chart-font graph operation failed: {0}")]
    Graph(#[source] Box<graph::Error>),
    /// A generated chart-style codec rejected the source.
    #[error("chart-font codec: {0}")]
    Codec(codec::Error),
    /// A shared wire operation rejected the source.
    #[error("chart-font wire: {0}")]
    Wire(#[source] litchi_iwa_common::Error),
    /// The package-wide object identifier space is exhausted.
    #[error("chart-font object identifier overflow")]
    IdentifierOverflow,
}

/// Bounded accounting for one chart-font operation. Scratch is a
/// conservative parser-buffer admission bound; retained bytes count
/// owned candidate/output buffers, while wire fields and output bytes are
/// operation totals.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Report {
    input_bytes: usize,
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    text_bytes: usize,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
    changed: bool,
}

impl Report {
    /// Whether the selected chart payload changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
    /// Source payload bytes inspected.
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }
    /// Candidate payload bytes emitted.
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }
    /// Wire fields visited.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }
    /// Aggregate graph and wire work.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }
    /// Maximum inheritance depth observed.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }
    /// Font-name bytes retained.
    #[must_use]
    pub const fn text_bytes(self) -> usize {
        self.text_bytes
    }
    /// Collection allocation units admitted.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }
    /// Source and candidate bytes retained.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }
    /// Scratch bytes reserved.
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }
}

#[derive(Debug, Clone, Copy)]
struct Budget {
    limits: Limits,
    input_bytes: usize,
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    text_bytes: usize,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

impl Budget {
    const fn new(limits: Limits) -> Self {
        Self {
            limits,
            input_bytes: 0,
            output_bytes: 0,
            fields: 0,
            work_bytes: 0,
            max_depth: 0,
            text_bytes: 0,
            allocations: 0,
            retained_bytes: 0,
            scratch_bytes: 0,
        }
    }

    fn limit(kind: LimitKind, observed: usize, maximum: usize) -> Error {
        Error::Limit {
            kind,
            observed: observed as u64,
            maximum: maximum as u64,
        }
    }

    fn add(
        &mut self,
        kind: LimitKind,
        current: usize,
        amount: usize,
        maximum: usize,
    ) -> Result<usize, Error> {
        let observed = current
            .checked_add(amount)
            .ok_or_else(|| Self::limit(kind, usize::MAX, maximum))?;
        if observed > maximum {
            return Err(Self::limit(kind, observed, maximum));
        }
        Ok(observed)
    }

    fn source(&mut self, bytes: usize) -> Result<(), Error> {
        self.input_bytes = self.add(
            LimitKind::InputBytes,
            self.input_bytes,
            bytes,
            self.limits.max_input_bytes(),
        )?;
        Ok(())
    }

    fn reserve_output(&mut self, bytes: usize) -> Result<(), Error> {
        if bytes > self.limits.max_output_bytes() {
            return Err(Self::limit(
                LimitKind::OutputBytes,
                bytes,
                self.limits.max_output_bytes(),
            ));
        }
        self.retained_bytes = self.add(
            LimitKind::RetainedBytes,
            self.retained_bytes,
            bytes,
            self.limits.max_retained_bytes(),
        )?;
        Ok(())
    }

    fn finish_output(&mut self, bytes: usize) {
        self.output_bytes = bytes;
    }

    fn fields(&mut self, fields: usize) -> Result<(), Error> {
        self.fields = self.add(
            LimitKind::Fields,
            self.fields,
            fields,
            self.limits.max_fields(),
        )?;
        Ok(())
    }

    fn work(&mut self, bytes: usize) -> Result<(), Error> {
        self.work_bytes = self.add(
            LimitKind::Work,
            self.work_bytes,
            bytes,
            self.limits.max_work_bytes(),
        )?;
        Ok(())
    }

    fn text(&mut self, bytes: usize) -> Result<(), Error> {
        self.text_bytes = self.add(
            LimitKind::TextBytes,
            self.text_bytes,
            bytes,
            self.limits.max_text_bytes(),
        )?;
        Ok(())
    }

    fn allocation(&mut self, units: usize) -> Result<(), Error> {
        self.allocations = self.add(
            LimitKind::Allocations,
            self.allocations,
            units,
            self.limits.max_allocations(),
        )?;
        Ok(())
    }

    fn scratch(&mut self, bytes: usize) -> Result<(), Error> {
        self.scratch_bytes = self.add(
            LimitKind::ScratchBytes,
            self.scratch_bytes,
            bytes,
            self.limits.max_scratch_bytes(),
        )?;
        Ok(())
    }

    fn wire_limits(self) -> Result<WireLimits, Error> {
        let input = self.remaining(
            LimitKind::InputBytes,
            self.limits.max_input_bytes(),
            self.input_bytes,
        )?;
        let fields = self.remaining(LimitKind::Fields, self.limits.max_fields(), self.fields)?;
        let output = self.remaining(
            LimitKind::OutputBytes,
            self.limits.max_output_bytes(),
            self.output_bytes,
        )?;
        let work = self.remaining(
            LimitKind::Work,
            self.limits.max_work_bytes(),
            self.work_bytes,
        )?;
        if self.max_depth > self.limits.max_nesting() {
            return Err(Self::limit(
                LimitKind::Nesting,
                self.max_depth as usize,
                self.limits.max_nesting() as usize,
            ));
        }
        if self.limits.max_nesting() == 0 {
            return Err(Self::limit(LimitKind::Nesting, 1, 0));
        }
        WireLimits::DEFAULT
            .with_input_bytes(input)
            .and_then(|limits| limits.with_fields(fields))
            .and_then(|limits| limits.with_output_bytes(output))
            .and_then(|limits| limits.with_nesting(self.limits.max_nesting() as usize))
            .and_then(|limits| limits.with_rewrite_work(work.min(WireLimits::MAX_REWRITE_WORK)))
            .map_err(map_wire_error)
    }

    fn codec_limits(self) -> Result<CodecLimits, Error> {
        let bytes = self.remaining(
            LimitKind::OutputBytes,
            self.limits.max_output_bytes(),
            self.output_bytes,
        )?;
        let fields = self.remaining(LimitKind::Fields, self.limits.max_fields(), self.fields)?;
        CodecLimits::DEFAULT
            .with_max_bytes(bytes)
            .and_then(|limits| limits.with_max_fields(fields))
            .map_err(map_codec_error)
    }

    fn registration_limits(self) -> Result<shape::RegistrationLimits, Error> {
        Ok(shape::RegistrationLimits {
            max_input_bytes: self.remaining(
                LimitKind::InputBytes,
                self.limits.max_input_bytes(),
                self.input_bytes,
            )?,
            max_fields: self.remaining(LimitKind::Fields, self.limits.max_fields(), self.fields)?,
            max_output_bytes: self.remaining(
                LimitKind::OutputBytes,
                self.limits.max_output_bytes(),
                self.output_bytes,
            )?,
            max_work_bytes: self.remaining(
                LimitKind::Work,
                self.limits.max_work_bytes(),
                self.work_bytes,
            )?,
            max_allocations: self.remaining(
                LimitKind::Allocations,
                self.limits.max_allocations(),
                self.allocations,
            )?,
            max_scratch_bytes: self.remaining(
                LimitKind::ScratchBytes,
                self.limits.max_scratch_bytes(),
                self.scratch_bytes,
            )?,
        })
    }

    fn charge_registration(&mut self, report: shape::RegistrationReport) -> Result<()> {
        self.source(report.input_bytes)?;
        self.fields(report.fields)?;
        self.work(report.work_bytes)?;
        self.allocation(report.allocations)?;
        self.scratch(report.scratch_bytes)?;
        self.reserve_output(report.output_bytes)?;
        self.finish_output(report.output_bytes);
        Ok(())
    }

    fn admit_registry(&mut self, workspace: &Workspace<'_>, extra: usize) -> Result<()> {
        let payload = workspace.registry_payload_len()?;
        if payload == 0 {
            return Ok(());
        }
        let retained = payload.saturating_add(extra);
        self.source(payload)?;
        self.fields(payload)?;
        self.work(payload.saturating_mul(4))?;
        self.allocation(4)?;
        self.scratch(retained)?;
        self.reserve_output(retained)?;
        Ok(())
    }

    fn remaining(self, kind: LimitKind, maximum: usize, used: usize) -> Result<usize, Error> {
        let remaining = maximum.saturating_sub(used);
        if remaining == 0 {
            return Err(Self::limit(kind, used.saturating_add(1), maximum));
        }
        Ok(remaining)
    }

    fn parse_fields(&mut self, source: &[u8]) -> Result<Vec<WireField>> {
        self.fields_at_depth(source, 1)
    }

    fn fields_at_depth(&mut self, source: &[u8], depth: u32) -> Result<Vec<WireField>> {
        if depth > self.limits.max_nesting() {
            return Err(Self::limit(
                LimitKind::Nesting,
                depth as usize,
                self.limits.max_nesting() as usize,
            ));
        }
        // The first pass is allocation-free; reserve only its fixed visitor
        // state before scanning. The exact field-vector capacity is admitted
        // after this pass and before the allocating second pass.
        self.scratch(size_of::<WireField>())?;
        self.allocation(1)?;
        let limits = self.wire_limits()?;
        self.source(source.len())?;
        self.work(source.len())?;
        let mut field_count = 0usize;
        wire::visit_wire_fields(source, limits, |_field| {
            field_count = field_count.saturating_add(1);
            Ok(())
        })
        .map_err(map_wire_error)?;
        self.fields(field_count)?;
        let field_storage = field_count
            .checked_mul(size_of::<WireField>())
            .ok_or_else(|| {
                Self::limit(
                    LimitKind::ScratchBytes,
                    usize::MAX,
                    self.limits.max_scratch_bytes(),
                )
            })?;
        self.scratch(field_storage)?;
        self.allocation(1)?;
        self.work(source.len().checked_add(field_count).ok_or_else(|| {
            Self::limit(LimitKind::Work, usize::MAX, self.limits.max_work_bytes())
        })?)?;
        let fields = wire::parse_wire_fields(source, limits).map_err(map_wire_error)?;
        debug_assert_eq!(fields.len(), field_count);
        // `parse_wire_fields` validates one message. Nested paragraph-style
        // messages are charged by their own call, so the observed depth is
        // the message level (root = one).
        self.max_depth = self.max_depth.max(depth);
        Ok(fields)
    }

    fn report(self, changed: bool) -> Report {
        Report {
            input_bytes: self.input_bytes,
            output_bytes: self.output_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            text_bytes: self.text_bytes,
            allocations: self.allocations,
            retained_bytes: self.retained_bytes,
            scratch_bytes: self.scratch_bytes,
            changed,
        }
    }
}

struct ParagraphStyles<'w, 's> {
    workspace: &'w Workspace<'s>,
    error: std::cell::RefCell<Option<Error>>,
    budget: std::cell::RefCell<Budget>,
}

type Result<T, E = Error> = std::result::Result<T, E>;

impl<'w, 's> ParagraphStyles<'w, 's> {
    fn new(workspace: &'w Workspace<'s>, budget: Budget) -> Self {
        Self {
            workspace,
            error: std::cell::RefCell::new(None),
            budget: std::cell::RefCell::new(budget),
        }
    }

    fn take_error(&self) -> Option<Error> {
        self.error.borrow_mut().take()
    }

    fn limits(&self) -> Result<WireLimits> {
        self.budget.borrow().wire_limits()
    }

    fn into_budget(self) -> Budget {
        self.budget.into_inner()
    }

    fn charge_style(&self, identifier: u64, payload: &[u8]) -> Result<()> {
        let mut budget = self.budget.borrow_mut();
        // The paragraph-style reader may materialize a validated font name;
        // admit one semantic value before delegating to the text-wire codec.
        budget.allocation(1)?;
        let wire_limits = budget.wire_limits()?;
        let mut fields = 0usize;
        wire::visit_wire_fields(payload, wire_limits, |_field| {
            fields = fields.saturating_add(1);
            Ok(())
        })
        .map_err(map_wire_error)?;
        budget.source(payload.len())?;
        budget.fields(fields)?;
        budget.work(fields)?;
        // The text-wire effective-property walk descends through the
        // paragraph-style envelope after the root payload has been checked.
        budget.max_depth = budget.max_depth.max(3);
        let _ = identifier;
        Ok(())
    }
}

impl Styles for ParagraphStyles<'_, '_> {
    fn paragraph_style(&self, identifier: u64) -> Option<&[u8]> {
        match self
            .workspace
            .unique_payload(identifier, PARAGRAPH_STYLE_MESSAGE_TYPE, "")
        {
            Ok((_, payload)) => {
                if let Err(error) = self.charge_style(identifier, payload) {
                    *self.error.borrow_mut() = Some(error);
                    None
                } else {
                    Some(payload)
                }
            },
            Err(error) => {
                *self.error.borrow_mut() = Some(map_graph_error(error));
                None
            },
        }
    }
}

fn map_wire_limit(kind: WireLimitKind) -> LimitKind {
    LimitKind::Wire(kind)
}

fn map_wire_error(error: litchi_iwa_common::Error) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => Error::Limit {
            kind: map_wire_limit(kind),
            observed: observed as u64,
            maximum: limit as u64,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation { amount },
        error => Error::Wire(error),
    }
}

fn map_codec_error(error: codec::Error) -> Error {
    match error {
        codec::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: LimitKind::Codec(kind),
            observed: observed as u64,
            maximum: maximum as u64,
        },
        codec::Error::Allocation { bytes } => Error::Allocation { amount: bytes },
        error => Error::Codec(error),
    }
}

fn map_text_error(error: litchi_iwa_text_wire::paragraph_style::Error) -> Error {
    use litchi_iwa_common::LimitKind as WireKind;
    match error {
        litchi_iwa_text_wire::paragraph_style::Error::InvalidSource { reason } => {
            Error::InvalidSource(reason)
        },
        litchi_iwa_text_wire::paragraph_style::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                WireKind::InputBytes => LimitKind::InputBytes,
                WireKind::Fields => LimitKind::Fields,
                WireKind::OutputBytes => LimitKind::OutputBytes,
                WireKind::Nesting => LimitKind::Nesting,
                WireKind::RewriteWork => LimitKind::Work,
                other => LimitKind::Wire(other),
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_text_wire::paragraph_style::Error::Allocation { amount } => {
            Error::Allocation { amount }
        },
        litchi_iwa_text_wire::paragraph_style::Error::Refused(reason) => Error::Refused(reason),
        litchi_iwa_text_wire::paragraph_style::Error::Internal => {
            Error::InvalidSource("paragraph-style read-back failed")
        },
        _ => Error::InvalidSource("unsupported paragraph-style failure"),
    }
}

fn map_graph_error(error: graph::Error) -> Error {
    match error {
        graph::Error::Archive(error) => match error {
            litchi_iwa_core::Error::Limit {
                kind,
                observed,
                maximum,
            } => Error::Limit {
                kind: LimitKind::Archive(kind),
                observed: observed as u64,
                maximum: maximum as u64,
            },
            litchi_iwa_core::Error::Allocation { requested, .. } => {
                Error::Allocation { amount: requested }
            },
            error => Error::Graph(Box::new(graph::Error::Archive(error))),
        },
        graph::Error::Wire(error) => map_wire_error(error),
        graph::Error::Codec(error) => map_codec_error(error),
        graph::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                graph::GraphLimitKind::ScratchBytes => LimitKind::ScratchBytes,
                graph::GraphLimitKind::AllocationEvents => LimitKind::Allocations,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        error => Error::Graph(Box::new(error)),
    }
}

fn map_registration_error(error: shape::RegistrationError) -> Error {
    match error {
        shape::RegistrationError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                shape::RegistrationLimitKind::InputBytes => LimitKind::InputBytes,
                shape::RegistrationLimitKind::Fields => LimitKind::Fields,
                shape::RegistrationLimitKind::OutputBytes => LimitKind::OutputBytes,
                shape::RegistrationLimitKind::Work => LimitKind::Work,
                shape::RegistrationLimitKind::Allocations => LimitKind::Allocations,
                shape::RegistrationLimitKind::ScratchBytes => LimitKind::ScratchBytes,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        shape::RegistrationError::Allocation { amount, .. } => Error::Allocation { amount },
        shape::RegistrationError::Wire(error) => map_wire_error(error),
        shape::RegistrationError::Graph(error) => map_graph_error(error),
        shape::RegistrationError::Invalid(_) => Error::InvalidSource("stylesheet registration"),
    }
}

impl From<graph::Error> for Error {
    fn from(error: graph::Error) -> Self {
        map_graph_error(error)
    }
}

impl From<litchi_iwa_common::Error> for Error {
    fn from(error: litchi_iwa_common::Error) -> Self {
        map_wire_error(error)
    }
}

impl From<codec::Error> for Error {
    fn from(error: codec::Error) -> Self {
        map_codec_error(error)
    }
}

impl From<litchi_iwa_text_wire::paragraph_style::Error> for Error {
    fn from(error: litchi_iwa_text_wire::paragraph_style::Error) -> Self {
        map_text_error(error)
    }
}

impl From<litchi_iwa_core::Error> for Error {
    fn from(error: litchi_iwa_core::Error) -> Self {
        map_graph_error(graph::Error::Archive(error))
    }
}

fn chart_message(
    workspace: &Workspace<'_>,
    chart: u64,
    budget: &mut Budget,
) -> Result<(usize, Vec<u8>)> {
    let (message, payload) = workspace
        .unique_payload(
            chart,
            CHART_MESSAGE_TYPE,
            "a chart must have exactly one chart payload",
        )
        .map_err(map_graph_error)?;
    let _ = budget.parse_fields(payload)?;
    budget.reserve_output(payload.len())?;
    budget.scratch(payload.len())?;
    budget.allocation(1)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(payload.len())
        .map_err(|_| Error::Allocation {
            amount: payload.len(),
        })?;
    output.extend_from_slice(payload);
    Ok((message, output))
}

fn paragraph_styles(
    workspace: &Workspace<'_>,
    chart: u64,
    budget: &mut Budget,
) -> Result<Vec<u64>> {
    let (_, payload) = chart_message(workspace, chart, budget)?;
    budget.allocation(1)?;
    budget.scratch(payload.len())?;
    let limits = budget.codec_limits()?;
    let references = codec::references(&payload, limits).map_err(map_codec_error)?;
    let source = references.paragraph_styles();
    budget.allocation(source.len().max(1))?;
    budget.scratch(source.len().saturating_mul(size_of::<u64>()))?;
    let mut styles = Vec::new();
    styles
        .try_reserve_exact(source.len())
        .map_err(|_| Error::Allocation {
            amount: source.len(),
        })?;
    styles.extend_from_slice(source);
    if styles.is_empty() {
        return Err(Error::InvalidSource("a chart has no paragraph styles"));
    }
    if styles.contains(&0) {
        return Err(Error::InvalidSource(
            "a chart has a zero paragraph-style reference",
        ));
    }
    Ok(styles)
}

fn unique(identifiers: &[u64], budget: &mut Budget) -> Result<Vec<u64>> {
    budget.allocation(identifiers.len().max(1))?;
    budget.scratch(identifiers.len().saturating_mul(size_of::<u64>()))?;
    let mut seen = HashSet::new();
    seen.try_reserve(identifiers.len())
        .map_err(|_| Error::Allocation {
            amount: identifiers.len(),
        })?;
    let mut unique = Vec::new();
    unique
        .try_reserve(identifiers.len())
        .map_err(|_| Error::Allocation {
            amount: identifiers.len(),
        })?;
    for identifier in identifiers {
        budget.work(size_of::<u64>())?;
        if seen.insert(*identifier) {
            unique.push(*identifier);
        }
    }
    Ok(unique)
}

fn effective_property(
    workspace: &Workspace<'_>,
    style: u64,
    kind: PropertyKind,
    budget: &mut Budget,
) -> Result<Property> {
    let styles = ParagraphStyles::new(workspace, *budget);
    let limits = styles.limits()?;
    let result = text_style::effective(&styles, style, kind, limits);
    let error = styles.take_error();
    *budget = styles.into_budget();
    if let Some(error) = error {
        return Err(error);
    }
    result.map_err(map_text_error)
}

fn effective_font(workspace: &Workspace<'_>, style: u64, budget: &mut Budget) -> Result<ChartFont> {
    let font = match effective_property(workspace, style, PropertyKind::TextFont, budget)? {
        Property::TextFont(font) => font,
        _ => return Err(Error::InvalidSource("an unexpected effective font")),
    };
    let style = text_style_of(workspace, style, budget)?;
    budget.text(font_name_bytes(&font))?;
    Ok(ChartFont {
        font,
        bold: style.bold,
        italic: style.italic,
    })
}

fn text_style_of(workspace: &Workspace<'_>, style: u64, budget: &mut Budget) -> Result<TextStyle> {
    match effective_property(workspace, style, PropertyKind::TextStyle, budget)? {
        Property::TextStyle(style) => Ok(style),
        _ => Err(Error::InvalidSource("an unexpected effective text style")),
    }
}

fn font_name_bytes(font: &Font) -> usize {
    match font {
        Font::Named(name) => name.as_str().len(),
        _ => 0,
    }
}

fn uniform<T: PartialEq>(
    workspace: &Workspace<'_>,
    chart: u64,
    budget: &mut Budget,
    read: impl Fn(&Workspace<'_>, u64, &mut Budget) -> Result<T>,
) -> Result<T> {
    let styles = paragraph_styles(workspace, chart, budget)?;
    let first = read(workspace, styles[0], budget)?;
    for style in &styles[1..] {
        if read(workspace, *style, budget)? != first {
            return Err(Error::InvalidSource("a chart has mixed effective fonts"));
        }
    }
    Ok(first)
}

/// The uniform effective font of chart drawable `chart`.
///
/// # Errors
///
/// An invalid chart or paragraph style, or mixed fonts.
pub fn font(workspace: &Workspace<'_>, chart: u64, limits: Limits) -> Result<(ChartFont, Report)> {
    let mut budget = Budget::new(limits);
    budget.work(workspace.source_work())?;
    let value = uniform(workspace, chart, &mut budget, effective_font)?;
    Ok((value, budget.report(false)))
}

/// The uniform effective point size of chart drawable `chart`.
///
/// # Errors
///
/// An invalid chart or paragraph style, or mixed sizes.
pub fn size(
    workspace: &Workspace<'_>,
    chart: u64,
    limits: Limits,
) -> Result<(TextPointSize, Report)> {
    let mut budget = Budget::new(limits);
    budget.work(workspace.source_work())?;
    let value = uniform(workspace, chart, &mut budget, |workspace, style, budget| {
        Ok(text_style_of(workspace, style, budget)?.point_size)
    })?;
    Ok((value, budget.report(false)))
}

/// `(parent, stylesheet)` of a paragraph-style payload.
fn style_links(payload: &[u8], budget: &mut Budget) -> Result<(Option<u64>, Option<u64>)> {
    let fields = budget.parse_fields(payload)?;
    let Some(base_field) = fields
        .iter()
        .copied()
        .find(|field| field.number() == STYLE_SUPER_FIELD && field.wire_type() == 2)
    else {
        return Err(Error::InvalidSource(
            "a paragraph style has no style parent",
        ));
    };
    let base = base_field.payload(payload).map_err(map_wire_error)?;
    let base_fields = budget.fields_at_depth(base, 2)?;
    let mut link = |number| -> Result<Option<u64>> {
        let Some(reference_field) = base_fields
            .iter()
            .copied()
            .rfind(|field| field.number() == number)
        else {
            return Ok(None);
        };
        if reference_field.wire_type() != 2 {
            return Err(Error::InvalidSource(
                "a paragraph-style link is not a reference",
            ));
        }
        let reference = reference_field.payload(base).map_err(map_wire_error)?;
        varint_field(reference, REFERENCE_IDENTIFIER_FIELD, budget, 3)
    };
    Ok((link(STYLE_PARENT_FIELD)?, link(STYLE_STYLESHEET_FIELD)?))
}

fn varint_field(
    payload: &[u8],
    number: u32,
    budget: &mut Budget,
    depth: u32,
) -> Result<Option<u64>> {
    let fields = budget.fields_at_depth(payload, depth)?;
    let Some(field) = fields
        .iter()
        .copied()
        .rfind(|field| field.number() == number)
    else {
        return Ok(None);
    };
    if field.wire_type() != 0 {
        return Err(Error::InvalidSource(
            "a reference identifier is not a varint",
        ));
    }
    let encoded = field.payload(payload).map_err(map_wire_error)?;
    let (value, width) = litchi_iwa_common::decode_varint_from_bytes(encoded)
        .map_err(|_| Error::InvalidSource("a reference identifier is malformed"))?;
    if width != encoded.len() {
        return Err(Error::InvalidSource(
            "a reference identifier is noncanonical",
        ));
    }
    Ok(Some(value))
}

fn style_payload<'w>(workspace: &'w Workspace<'_>, style: u64) -> Result<(usize, &'w [u8])> {
    workspace
        .unique_payload(
            style,
            PARAGRAPH_STYLE_MESSAGE_TYPE,
            "a paragraph style must have exactly one payload",
        )
        .map_err(map_graph_error)
}

struct Addition {
    identifier: u64,
    parent: u64,
    stylesheet: u64,
    variation: text_style::Variation,
}

fn font_only(
    workspace: &Workspace<'_>,
    style: u64,
    budget: &mut Budget,
) -> Result<Option<FontOverrides>> {
    let styles = ParagraphStyles::new(workspace, *budget);
    let limits = styles.limits()?;
    let result = text_style::font_overrides(&styles, style, limits);
    let error = styles.take_error();
    *budget = styles.into_budget();
    if let Some(error) = error {
        return Err(error);
    }
    result.map_err(map_text_error)
}

fn apply(
    overrides: &mut FontOverrides,
    inherited_font: &ChartFont,
    inherited_size: TextPointSize,
    change: &Change,
    budget: &mut Budget,
) -> Result<()> {
    match change {
        Change::Font(target) => {
            budget.allocation(1)?;
            budget.text(font_name_bytes(&target.font))?;
            overrides.font = (inherited_font.font != target.font).then(|| target.font.clone());
            overrides.bold = (inherited_font.bold != target.bold).then_some(target.bold);
            overrides.italic = (inherited_font.italic != target.italic).then_some(target.italic);
        },
        Change::Size(target) => {
            overrides.point_size = (inherited_size != *target).then_some(*target);
        },
        Change::ResetFont | Change::ResetSize => {},
    }
    Ok(())
}

fn font_variation(
    parent: u64,
    stylesheet: u64,
    overrides: &FontOverrides,
    budget: &mut Budget,
) -> Result<text_style::Variation> {
    // The paragraph-style encoder admits its output before allocating. The
    // estimate covers the fixed envelope and the caller's font name; the
    // codec then applies the stricter wire limits to the exact candidate.
    let estimate =
        font_name_bytes(overrides.font.as_ref().unwrap_or(&Font::default())).saturating_add(256);
    budget.allocation(1)?;
    budget.scratch(estimate)?;
    let limits = budget.wire_limits()?;
    text_style::font_variation(parent, stylesheet, overrides, limits).map_err(map_text_error)
}

/// Apply `change` to chart drawable `chart`. Returns `false` when nothing
/// changes (the value is already the requested one, or nothing to reset).
///
/// # Errors
///
/// An invalid chart, paragraph style, stylesheet or registry; a write that
/// does not read back.
pub fn set(
    workspace: &mut Workspace<'_>,
    chart: u64,
    change: &Change,
    limits: Limits,
) -> Result<Report> {
    let mut budget = Budget::new(limits);
    budget.work(workspace.source_work())?;
    match change {
        Change::Font(target) => {
            let current = uniform(workspace, chart, &mut budget, effective_font)?;
            if current == *target {
                return Ok(budget.report(false));
            }
        },
        Change::Size(target) => {
            let current = uniform(workspace, chart, &mut budget, |workspace, style, budget| {
                Ok(text_style_of(workspace, style, budget)?.point_size)
            })?;
            if current == *target {
                return Ok(budget.report(false));
            }
        },
        _ => {},
    }
    let original = workspace;
    budget.allocation(2)?;
    budget.scratch(original.source_work())?;
    let mut staged = original.fork();
    let workspace = &mut staged;
    let references = paragraph_styles(workspace, chart, &mut budget)?;
    let unique = unique(&references, &mut budget)?;
    budget.allocation(unique.len().saturating_mul(2).max(1))?;
    budget.scratch(unique.len().saturating_mul(64))?;
    let mut next = workspace.next_object_identifier()?;
    let mut replacements = BTreeMap::new();
    let mut additions = Vec::new();
    additions
        .try_reserve(unique.len())
        .map_err(|_| Error::Allocation {
            amount: unique.len(),
        })?;
    let mut candidates = 0usize;
    for &style in &unique {
        let (_, payload) = style_payload(workspace, style)?;
        let (parent, stylesheet) = style_links(payload, &mut budget)?;
        let stylesheet =
            stylesheet.ok_or(Error::InvalidSource("a paragraph style has no stylesheet"))?;
        let direct = font_only(workspace, style, &mut budget)?;
        let (base, mut overrides) = match change {
            Change::Font(_) | Change::Size(_) => match direct {
                Some(overrides) => (
                    parent.ok_or(Error::InvalidSource(
                        "a paragraph-style variation has no parent",
                    ))?,
                    overrides,
                ),
                None => (style, FontOverrides::default()),
            },
            Change::ResetFont | Change::ResetSize => {
                let Some(mut overrides) = direct else {
                    continue;
                };
                let present = if *change == Change::ResetFont {
                    let present = overrides.font.is_some()
                        || overrides.bold.is_some()
                        || overrides.italic.is_some();
                    overrides.font = None;
                    overrides.bold = None;
                    overrides.italic = None;
                    present
                } else {
                    overrides.point_size.take().is_some()
                };
                if !present {
                    continue;
                }
                (
                    parent.ok_or(Error::InvalidSource(
                        "a paragraph-style variation has no parent",
                    ))?,
                    overrides,
                )
            },
        };
        candidates += 1;
        if matches!(change, Change::Font(_) | Change::Size(_)) {
            let inherited_font = effective_font(workspace, base, &mut budget)?;
            let inherited_size = text_style_of(workspace, base, &mut budget)?.point_size;
            apply(
                &mut overrides,
                &inherited_font,
                inherited_size,
                change,
                &mut budget,
            )?;
        }
        if overrides.is_empty() {
            replacements.insert(style, base);
            continue;
        }
        let identifier = next;
        next = next.checked_add(1).ok_or(Error::IdentifierOverflow)?;
        let variation = font_variation(base, stylesheet, &overrides, &mut budget)?;
        replacements.insert(style, identifier);
        additions.push(Addition {
            identifier,
            parent: base,
            stylesheet,
            variation,
        });
    }
    if candidates == 0 {
        return Ok(budget.report(false));
    }
    let chart_location = workspace.locate(chart)?;
    let chart_member_len = workspace.member(chart_location).len();
    budget.allocation(1)?;
    budget.scratch(chart_member_len)?;
    let chart_member = workspace.member(chart_location).to_owned();
    repoint(workspace, chart, &replacements, &mut budget)?;
    let last = additions.last().map(|addition| addition.identifier);
    for addition in additions {
        insert(workspace, &chart_member, addition, &mut budget)?;
    }
    if let Some(last) = last {
        budget.admit_registry(workspace, 64)?;
        workspace.set_last_object_identifier(last)?;
    }
    prune(workspace, &chart_member, &unique, &mut budget)?;
    let result = match change {
        Change::Font(target)
            if uniform(workspace, chart, &mut budget, effective_font)? != *target =>
        {
            Err(Error::InvalidSource(
                "a chart-font write does not read back",
            ))
        },
        Change::Size(target)
            if uniform(workspace, chart, &mut budget, |workspace, style, budget| {
                Ok(text_style_of(workspace, style, budget)?.point_size)
            })? != *target =>
        {
            Err(Error::InvalidSource(
                "a chart font-size write does not read back",
            ))
        },
        _ => Ok(budget.report(true)),
    };
    match result {
        Ok(report) => {
            *original = staged;
            Ok(report)
        },
        Err(error) => Err(error),
    }
}

fn repoint(
    workspace: &mut Workspace<'_>,
    chart: u64,
    replacements: &BTreeMap<u64, u64>,
    budget: &mut Budget,
) -> Result<()> {
    let (message, payload) = chart_message(workspace, chart, budget)?;
    let projected = payload
        .len()
        .saturating_add(replacements.len().saturating_mul(16));
    budget.reserve_output(projected)?;
    budget.allocation(replacements.len().saturating_add(1))?;
    budget.scratch(projected)?;
    let data = wire::transform_length_delimited_fields_at_path(
        &payload,
        &[CHART_EXTENSION_FIELD, CHART_PARAGRAPH_STYLES_FIELD],
        |reference| -> Result<Vec<u8>> {
            let identifier =
                varint_field(reference, REFERENCE_IDENTIFIER_FIELD, budget, 3)?.unwrap_or(0);
            match replacements.get(&identifier) {
                Some(replacement) => Ok(wire::patch_varint_field(
                    reference,
                    REFERENCE_IDENTIFIER_FIELD,
                    true,
                    Some(*replacement),
                )
                .map_err(map_wire_error)?),
                None => {
                    budget.allocation(1)?;
                    budget.scratch(reference.len())?;
                    let mut copy = Vec::new();
                    copy.try_reserve_exact(reference.len())
                        .map_err(|_| Error::Allocation {
                            amount: reference.len(),
                        })?;
                    copy.extend_from_slice(reference);
                    Ok(copy)
                },
            }
        },
    )?;
    budget.finish_output(data.len());
    workspace.replace_payload(chart, message, data)?;
    let next = paragraph_styles(workspace, chart, budget)?;
    let references = &mut workspace
        .object_mut(chart)?
        .archive_info
        .message_infos
        .get_mut(message)
        .ok_or(Error::InvalidSource("a chart payload has no metadata"))?
        .object_references;
    references.retain(|identifier| !replacements.contains_key(identifier));
    for identifier in unique(&next, budget)? {
        if !references.contains(&identifier) {
            references.push(identifier);
        }
    }
    Ok(())
}

fn insert(
    workspace: &mut Workspace<'_>,
    chart_member: &str,
    addition: Addition,
    budget: &mut Budget,
) -> Result<()> {
    let (message, sheet) = workspace.unique_payload(
        addition.stylesheet,
        STYLESHEET_MESSAGE_TYPE,
        "a stylesheet must have exactly one payload",
    )?;
    let registration_limits = budget.registration_limits()?;
    let (data, registration_report) = shape::stylesheet_with_variation_admitted(
        sheet,
        addition.parent,
        addition.identifier,
        registration_limits,
    )
    .map_err(map_registration_error)?;
    budget.charge_registration(registration_report)?;
    workspace.replace_payload(addition.stylesheet, message, data)?;
    let info = workspace
        .object_mut(addition.stylesheet)?
        .archive_info
        .message_infos
        .get_mut(message)
        .ok_or(Error::InvalidSource("a stylesheet payload has no metadata"))?;
    if info.object_references.contains(&addition.identifier) {
        return Err(Error::InvalidSource(
            "the stylesheet metadata already references a new style",
        ));
    }
    info.object_references.push(addition.identifier);
    let reference_count = addition.variation.references().len();
    budget.allocation(reference_count.saturating_add(1))?;
    budget.scratch(reference_count.saturating_mul(size_of::<u64>()))?;
    let mut references = Vec::new();
    references
        .try_reserve_exact(reference_count)
        .map_err(|_| Error::Allocation {
            amount: reference_count,
        })?;
    references.extend_from_slice(addition.variation.references());
    let variation_payload = addition.variation.into_payload();
    budget.reserve_output(variation_payload.len().saturating_add(64))?;
    budget.allocation(1)?;
    budget.scratch(variation_payload.len().saturating_add(64))?;
    let mut object = ArchiveObject::new(
        addition.identifier,
        vec![RawMessage {
            type_: PARAGRAPH_STYLE_MESSAGE_TYPE,
            data: variation_payload,
        }],
        workspace.limits(),
    )?;
    let info = &mut object.archive_info.message_infos[0];
    info.versions = text_style::MESSAGE_VERSION.to_vec().into();
    info.object_references.extend(references);
    let style_location = workspace.locate(addition.stylesheet)?;
    let style_member_len = workspace.member(style_location).len();
    budget.allocation(1)?;
    budget.scratch(style_member_len)?;
    workspace.insert(style_location, object)?;
    let style_member = workspace.member(style_location).to_owned();
    budget.work(workspace.registry_inspection_work())?;
    if let Some(style_component) = workspace.component_for_member(&style_member)? {
        budget.admit_registry(workspace, 64)?;
        workspace.add_object_uuids(style_component, &[addition.identifier])?;
        budget.work(workspace.registry_inspection_work())?;
        if let Some(owner) = workspace.component_for_member(chart_member)?
            && owner != style_component
        {
            budget.admit_registry(workspace, 64)?;
            workspace.add_external_reference(owner, style_component, addition.identifier)?;
        }
    }
    Ok(())
}

/// Whether any object other than its stylesheet or child styles references
/// `style` in its metadata.
fn has_consumers(workspace: &Workspace<'_>, style: u64, budget: &mut Budget) -> Result<bool> {
    for location in workspace.locations_iter() {
        budget.work(size_of::<u64>())?;
        for object in &workspace.archive_at(location)?.objects {
            budget.work(
                object
                    .messages
                    .len()
                    .saturating_mul(size_of::<u64>())
                    .saturating_add(
                        object
                            .archive_info
                            .message_infos
                            .iter()
                            .map(|info| info.object_references.len())
                            .sum::<usize>(),
                    ),
            )?;
            for (message, info) in object
                .messages
                .iter()
                .zip(&object.archive_info.message_infos)
            {
                if !info.object_references.contains(&style)
                    || message.type_ == STYLESHEET_MESSAGE_TYPE
                {
                    continue;
                }
                if message.type_ == PARAGRAPH_STYLE_MESSAGE_TYPE
                    && style_links(&message.data, budget)?.0 == Some(style)
                {
                    continue;
                }
                return Ok(true);
            }
        }
    }
    Ok(false)
}

fn owner_reference_if_unused(
    workspace: &mut Workspace<'_>,
    chart_member: &str,
    style_member: &str,
    style: u64,
    budget: &mut Budget,
) -> Result<()> {
    budget.work(workspace.registry_inspection_work())?;
    let Some(style_component) = workspace.component_for_member(style_member)? else {
        return Ok(());
    };
    budget.work(workspace.registry_inspection_work())?;
    let Some(owner) = workspace.component_for_member(chart_member)? else {
        return Ok(());
    };
    if owner == style_component {
        return Ok(());
    }
    for location in workspace.locations_iter() {
        let member = workspace.member(location);
        budget.work(size_of::<u64>().saturating_add(member.len()))?;
        budget.allocation(1)?;
        budget.scratch(member.len())?;
        let member = member.to_owned();
        budget.work(size_of::<u64>())?;
        budget.work(workspace.registry_inspection_work())?;
        if workspace.component_for_member(&member)? != Some(owner) {
            continue;
        }
        let archive = workspace.archive_at(location)?;
        let mut found = false;
        for object in &archive.objects {
            let references = object
                .archive_info
                .message_infos
                .iter()
                .map(|info| info.object_references.len())
                .sum::<usize>();
            budget.work(references)?;
            if object
                .archive_info
                .message_infos
                .iter()
                .any(|info| info.object_references.contains(&style))
            {
                found = true;
                break;
            }
        }
        if found {
            return Ok(());
        }
    }
    budget.admit_registry(workspace, 64)?;
    Ok(workspace.remove_external_reference(owner, style_component, style)?)
}

fn prune(
    workspace: &mut Workspace<'_>,
    chart_member: &str,
    previous: &[u64],
    budget: &mut Budget,
) -> Result<()> {
    budget.allocation(previous.len().max(1))?;
    budget.scratch(previous.len().saturating_mul(size_of::<u64>()))?;
    let mut removed = Vec::new();
    removed
        .try_reserve(previous.len())
        .map_err(|_| Error::Allocation {
            amount: previous.len(),
        })?;
    for &style in previous {
        budget.work(size_of::<u64>())?;
        let location = workspace.locate(style)?;
        let style_member_len = workspace.member(location).len();
        budget.allocation(1)?;
        budget.scratch(style_member_len)?;
        let style_member = workspace.member(location).to_owned();
        if font_only(workspace, style, budget)?.is_none()
            || has_consumers(workspace, style, budget)?
        {
            owner_reference_if_unused(workspace, chart_member, &style_member, style, budget)?;
            continue;
        }
        let (_, payload) = style_payload(workspace, style)?;
        let (parent, stylesheet) = style_links(payload, budget)?;
        let parent = parent.ok_or(Error::InvalidSource(
            "a paragraph-style variation has no parent",
        ))?;
        let stylesheet =
            stylesheet.ok_or(Error::InvalidSource("a paragraph style has no stylesheet"))?;
        let (message, sheet) = workspace.unique_payload(
            stylesheet,
            STYLESHEET_MESSAGE_TYPE,
            "a stylesheet must have exactly one payload",
        )?;
        let registration_limits = budget.registration_limits()?;
        let (data, registration_report) =
            shape::stylesheet_without_variation_admitted(sheet, parent, style, registration_limits)
                .map_err(map_registration_error)?;
        budget.charge_registration(registration_report)?;
        workspace.replace_payload(stylesheet, message, data)?;
        let info = workspace
            .object_mut(stylesheet)?
            .archive_info
            .message_infos
            .get_mut(message)
            .ok_or(Error::InvalidSource("a stylesheet payload has no metadata"))?;
        if info
            .object_references
            .iter()
            .filter(|id| **id == style)
            .count()
            != 1
        {
            return Err(Error::InvalidSource(
                "the stylesheet metadata must reference a removed style once",
            ));
        }
        info.object_references.retain(|id| *id != style);
        for field in &mut info.field_infos {
            field.object_references.retain(|id| *id != style);
        }
        workspace.remove(style)?;
        budget.admit_registry(workspace, 64)?;
        budget.work(workspace.registry_inspection_work())?;
        if let Some(style_component) = workspace.component_for_member(&style_member)? {
            budget.admit_registry(workspace, 64)?;
            workspace.remove_object_uuids(style_component, &[style])?;
            budget.admit_registry(workspace, 64)?;
            workspace.remove_external_references_to(style_component, style)?;
            budget.work(workspace.registry_inspection_work())?;
            if let Some(owner) = workspace.component_for_member(chart_member)?
                && owner != style_component
            {
                budget.admit_registry(workspace, 64)?;
                workspace.add_external_reference(owner, style_component, parent)?;
            }
        }
        removed.push(style);
    }
    if !removed.is_empty() {
        workspace.release_identifier_suffix(&removed)?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::{Budget, Error, LimitKind, Limits, map_text_error};
    use litchi_iwa_text_wire::paragraph_style::{Error as TextError, Refusal};

    fn generous() -> Limits {
        Limits::new(1024, 128, 1024, 8, 4096, 1024, 128, 4096, 4096)
    }

    #[test]
    fn residual_zero_is_a_typed_refusal() {
        let mut budget = Budget::new(Limits::new(1, 16, 16, 4, 64, 16, 16, 64, 64));
        budget.source(1).expect("source admission");
        assert!(matches!(
            budget.wire_limits(),
            Err(Error::Limit {
                kind: LimitKind::InputBytes,
                observed: 2,
                maximum: 1,
            })
        ));
    }

    #[test]
    fn fields_are_counted_as_fields_and_accumulate() {
        let mut budget = Budget::new(generous());
        let first = [0x08, 0x01];
        let second = [0x10, 0x02, 0x18, 0x03];
        budget.parse_fields(&first).expect("first wire message");
        budget.parse_fields(&second).expect("second wire message");
        assert_eq!(budget.report(false).fields(), 3);
        assert_eq!(budget.report(false).input_bytes(), 6);
    }

    #[test]
    fn nested_depth_is_retained_in_the_report() {
        let mut budget = Budget::new(generous());
        budget
            .fields_at_depth(&[0x08, 0x01], 3)
            .expect("bounded nested message");
        assert_eq!(budget.report(false).max_depth(), 3);
    }

    #[test]
    fn zero_scratch_and_allocation_ceilings_fail_before_wire_work() {
        let mut scratch = Budget::new(Limits::new(32, 32, 32, 4, 128, 32, 32, 128, 0));
        assert!(matches!(
            scratch.parse_fields(&[0x08, 0x01]),
            Err(Error::Limit {
                kind: LimitKind::ScratchBytes,
                ..
            })
        ));

        let mut allocations = Budget::new(Limits::new(32, 32, 32, 4, 128, 32, 0, 128, 128));
        assert!(matches!(
            allocations.parse_fields(&[0x08, 0x01]),
            Err(Error::Limit {
                kind: LimitKind::Allocations,
                ..
            })
        ));
    }

    #[test]
    fn paragraph_style_refusals_remain_typed() {
        let error = map_text_error(TextError::Refused(Refusal::NotNamed));
        assert!(matches!(error, Error::Refused(Refusal::NotNamed)));
    }
}
