//! Public semantic transaction and worksheet-editing facade.

mod conflicts;
mod guard;
mod order;
mod snapshot;

use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet, btree_map::Entry};
use std::mem::size_of;
use std::sync::Arc;

use litchi_ooxml_common::mce::{Capabilities, StreamError, StreamLimits};
use litchi_ooxml_common::web as common_web;
use litchi_opc::{Part, Relationship, Relationships, TargetMode};
use litchi_sheet::{
    Area, At, Cell as Address, Column as ColumnIndex, ColumnAt, Rect, Row as RowIndex, RowAt,
};

use super::super::super::model::ValidatedWorksheetStore;
use super::super::super::{Selector, Visibility, Workbook, Worksheet, WorksheetKind};
use crate::Style;
use crate::cell::{Cell, Content};
use crate::column::{OutlineAt, State as ColumnState, WidthAt};
use crate::error::{EditBlock, Error, RemoveBlock, Result, TabEditBlock, allocation, invalid};
use crate::form_control::OrdinaryFormControlOverlay;
use crate::formula::{Formula, Kind as FormulaKind};
use crate::layout;
use crate::raw;
use crate::raw::worksheet::edit::{
    Action, ColumnAction, DefaultsAction, DescentEffect, HeightEffect, MergePlan, OptionalEffect,
    Payload, Plan, RowAction, StyleEffect, WidthEffect,
};
use crate::row::{HeightAt, State as RowState};
use crate::sheet::Name;
use crate::style::StyleLineage;
use crate::style::cell_format::Format;
use crate::style::{AlignmentChange, BorderChange, FillChange, FontChange, ProtectionChange};
use crate::web::{Binding as WebBinding, Bindings as WebBindings};

use super::super::model::{
    ContentTypesChange, PartChange, PatchAuthority, RelationshipChange, StyleGuard, SvgPartChange,
    defaults_after, ensure_merge_area, merge_conflicts, project_merges,
    validate_svg_final_removals,
};
use super::super::svg_lifecycle as svg_plan;
use super::super::validation::{
    Added, FinalOrder, HyperlinkAction, MergeIntent, OptionalAction, OrderPlan, PanesAction,
    Placement, SheetActions, TabAction, Target, pending_merge,
};
use super::super::{
    ActiveTab, Change, Commit, Conflict, ConflictSet, JoinError, JoinFailure, JoinNormalization,
    PackageChange, Patch, State,
};
use super::super::{codec, package, workbook_retained_bytes};

use self::snapshot::Snapshot;
use super::super::svg::{PictureSelector, SvgInput};
use super::worksheet::{
    NewSheet, Sheet, Tab, alignment_change_matches_source, border_change_matches_source,
    fill_change_matches_source, font_change_matches_source, number_format_matches_source,
    protection_change_matches_source, quote_prefix_matches_source,
};

const MAX_SVG_INPUT_BYTES: usize = 32 * 1024 * 1024;

#[derive(Debug, Clone)]
pub(crate) struct SvgLifecycleIntent {
    pub(crate) position: usize,
    pub(crate) selector: PictureSelector,
    pub(crate) attach: bool,
    pub(crate) payload: Option<Arc<Vec<u8>>>,
}

const MAX_CELL_TRANSFER: u64 = 65_536;
const MAX_STYLE_REQUESTS: usize = 65_536;
const MAX_STYLE_NUMBER_FORMATS: usize = 65_534;
const MAX_STYLE_FONTS: usize = 65_534;
const MAX_STYLE_FILLS: usize = 65_534;
const MAX_STYLE_BORDERS: usize = 65_534;
const MAX_STYLE_CELL_XFS: usize = 65_430;
const BORDER_STAGED_COPY_HEADROOM: usize = 3;
const BORDER_STAGED_NODE_BYTES: usize = 64;
const MAX_CELL_DEPENDENCY_SCAN: usize = 1_048_576;
const MAX_VALIDATED_STORE_HANDOFF_CELLS: usize = 4_096;
const MAX_VALIDATED_STORE_HANDOFF_BYTES: usize = 1_048_576;

/// The post-write readback that verifies one changed worksheet.
#[derive(Debug)]
enum VerifiedStore {
    /// The complete published worksheet was parsed.
    Complete(crate::cell::Store),
    /// Only the worksheet envelope, every row shell and every changed cell
    /// were parsed; every omitted cell record is a verbatim source record.
    Reduced(crate::cell::Store),
}

impl VerifiedStore {
    fn store(&self) -> &crate::cell::Store {
        match self {
            Self::Complete(store) | Self::Reduced(store) => store,
        }
    }
}

/// Whether a value-only rewrite may verify through the reduced readback.
///
/// Change 0525 admitted this readback for the source-backed route. The
/// eager route admits it only where the argument is local to cell records:
///
/// * the rewrite recorded which cell spans it copied verbatim from the
///   source, whose complete parse already validated them (with their styles)
///   in the same row, namespace and shared-string context;
/// * compaction left the rewrite unchanged, so those spans index the bytes
///   that will be published;
/// * the published `<sheetData>` body lies wholly in the benign lane subset
///   (rows of cells with at most one plain value: no formula, inline string,
///   foreign or markup-compatibility content), and the part does not mention
///   the markup-compatibility namespace, so no preprocessing can relate an
///   omitted record to a retained one;
/// * the complete store would exceed the validated-store handoff bounds, so
///   the reduced store never reaches a published snapshot;
/// * the worksheet fits the web-extension reader's limit. Every larger
///   changed worksheet is refused at that later step anyway, and below it the
///   complete parse's whole-document input limits cannot apply, so no first
///   error can move.
///
/// Anything else, and any refusal of the reduced readback itself, takes the
/// complete parse, whose result and error remain authoritative.
fn reduced_readback_admitted(
    after: &[u8],
    compacted: &raw::compact::WorksheetOutput,
    omitted: &[raw::worksheet::edit::OmittedCells],
    source_cells: usize,
    changed_cells: usize,
) -> bool {
    !omitted.is_empty()
        && after.len() <= raw::web::MAX_XML_BYTES
        && compacted.body_admitted()
        && compacted.bytes() == after
        && (after.len() > MAX_VALIDATED_STORE_HANDOFF_BYTES
            || source_cells.saturating_sub(changed_cells) > MAX_VALIDATED_STORE_HANDOFF_CELLS)
        && !litchi_ooxml_common::mce::mentions_namespace(after)
}

/// Verify a changed worksheet through the reduced readback, or `None` when
/// the complete parse must verify it.
///
/// Besides [`reduced_readback_admitted`], the reduced store is refused when
/// one of its cells falls inside an omitted range, as change 0525's merge
/// refuses it: such a record would sit beside the omitted verbatim record of
/// the same address, a duplicate only the complete parse would see.
fn reduced_readback_store<'s>(
    after: &[u8],
    compacted: &raw::compact::WorksheetOutput,
    omitted: &[raw::worksheet::edit::OmittedCells],
    source_cells: usize,
    changed_cells: usize,
    strings: impl FnOnce() -> Result<Option<&'s [crate::cell::Text]>>,
) -> Option<crate::cell::Store> {
    if !reduced_readback_admitted(after, compacted, omitted, source_cells, changed_cells) {
        return None;
    }
    let mut ranges = Vec::new();
    ranges.try_reserve_exact(omitted.len()).ok()?;
    for span in omitted {
        ranges.push(span.range()?);
    }
    // The speculative buffer is dropped before any complete-parse fallback.
    let reduced = raw::worksheet::edit::reduced_readback(after, omitted).ok()?;
    let store = raw::worksheet::parse(&reduced, strings).ok()?;
    store.avoids_omitted_cells(&ranges).then_some(store)
}
const MAX_HYPERLINK_EDITS: usize = 4_096;
const MAX_SPARSE_VERIFICATION_BYTES: usize = 16 * 1024 * 1024;

#[derive(Clone, Copy)]
enum CellTransfer {
    Copy,
    Move,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
struct TransferStyles {
    font: Option<FontChange>,
    alignment: Option<AlignmentChange>,
    fill: Option<FillChange>,
    border: Option<BorderChange>,
    protection: Option<ProtectionChange>,
    quote_prefix: Option<Option<bool>>,
    number_format: Option<Option<Arc<str>>>,
}

type StagedTransfer = (Address, Address, Action, TransferStyles, bool, bool);

impl TransferStyles {
    const fn is_empty(&self) -> bool {
        self.font.is_none()
            && self.alignment.is_none()
            && self.fill.is_none()
            && self.border.is_none()
            && self.protection.is_none()
            && self.quote_prefix.is_none()
            && self.number_format.is_none()
    }
}

/// Verify a large worksheet's sparse cell rewrite without materializing every
/// unchanged cell a second time.
///
/// Ok(false) means the worksheet is outside the scanner's first semantic
/// slice, so the caller must retain the full materialized-parser fallback.
fn verify_sparse_changed_cells(
    content: &[u8],
    targets: &[Address],
    changes: &[Change],
) -> Result<bool> {
    if content.is_empty() || content.len() > MAX_SPARSE_VERIFICATION_BYTES {
        return Ok(false);
    }
    if changes
        .iter()
        .any(|change| !matches!(change, Change::Cell { .. }))
    {
        return Ok(false);
    }
    if changes.iter().any(|change| {
        matches!(
            change,
            Change::Cell {
                after: State::Cell {
                    style: crate::StyleState::Shared(_),
                    ..
                },
                ..
            }
        )
    }) {
        return Ok(false);
    }

    let mut limits = StreamLimits::default();
    limits.processing.max_input_bytes = content.len();
    limits.processing.max_output_bytes = content.len();
    let outcome = raw::worksheet::selected::scan_targets_bytes(
        content,
        &Capabilities::default(),
        &limits,
        targets,
    )
    .map_err(map_sparse_stream_error)?;
    let selected = match outcome {
        raw::selected_worksheet::RangeScanOutcome::Eligible(selected) => selected,
        raw::selected_worksheet::RangeScanOutcome::NotEligible(_) => return Ok(false),
    };
    // Shared strings and direct styles need workbook-level resolution. The
    // source store was already parsed, but this sparse verifier intentionally
    // does not rebuild those dependency tables; retain the existing full-
    // parser fallback whenever the output still contains one.
    if selected.dependencies.max_shared_string_index.is_some()
        || selected.dependencies.max_direct_style_index.is_some()
    {
        return Ok(false);
    }

    for change in changes {
        let Change::Cell { address, after, .. } = change else {
            return Ok(false);
        };
        let record = selected
            .cells
            .binary_search_by_key(address, |record| record.address)
            .ok()
            .and_then(|index| selected.cells.get(index));
        match after {
            State::Missing => {
                if record.is_some() {
                    return Err(invalid(format!(
                        "sparse worksheet verification retained removed cell {address}"
                    )));
                }
            },
            State::Cell {
                content: expected,
                style,
                shared_string,
            } => {
                if !matches!(style, crate::StyleState::Default) || shared_string.is_some() {
                    return Ok(false);
                }
                let record = record.ok_or_else(|| {
                    invalid(format!(
                        "sparse worksheet verification lost changed cell {address}"
                    ))
                })?;
                // A deferred shared-string payload is a mismatch: only a
                // validated semantic cell equal to the staged content passes.
                if !matches!(
                    &record.payload,
                    raw::selected_worksheet::SelectedPayload::Cell(cell) if cell == expected
                ) {
                    return Err(invalid(format!(
                        "sparse worksheet edit verification failed at {address}"
                    )));
                }
            },
        }
    }
    Ok(true)
}

fn map_sparse_stream_error(error: StreamError<Error, Error>) -> Error {
    match error {
        StreamError::Input {
            raw_error: Some(error),
            ..
        }
        | StreamError::Mce {
            raw_error: Some(error),
            ..
        }
        | StreamError::Callback {
            raw_error: Some(error),
            ..
        } => error,
        StreamError::Callback {
            raw_error: None,
            active_error: Some(error),
        }
        | StreamError::Input {
            raw_error: None,
            active_error: Some(error),
            ..
        }
        | StreamError::Mce {
            raw_error: None,
            active_error: Some(error),
            ..
        } => error,
        StreamError::Input { error, .. } => Error::Package(litchi_opc::OpcError::IoError(error)),
        StreamError::Mce { error, .. } => Error::MarkupCompatibility(error),
        StreamError::Callback {
            raw_error: None,
            active_error: None,
        } => invalid("sparse worksheet stream failed without a callback error"),
        _ => invalid("unknown sparse worksheet stream error"),
    }
}

/// Isolated workbook transaction. Dropping it rolls back every pending change.
#[derive(Debug, Clone)]
pub struct Edit {
    pub(in crate::workbook::edit) base: Workbook,
    pub(in crate::workbook::edit) panes: Option<PanesAction>,
    pub(in crate::workbook::edit) defined_names: Option<Box<[raw::DefinedName]>>,
    pub(in crate::workbook::edit) drawings: BTreeMap<usize, super::super::drawing_transfer::Plan>,
    pub(in crate::workbook::edit) active: Option<Target>,
    pub(in crate::workbook::edit) order: Option<OrderPlan>,
    pub(in crate::workbook::edit) sheets: BTreeMap<usize, SheetActions>,
    pub(in crate::workbook::edit) added: Vec<Added>,
    pub(in crate::workbook::edit) removed: BTreeSet<usize>,
    pub(in crate::workbook::edit) cross_workbook_scalar: bool,
    pub(in crate::workbook::edit) svg_lifecycle: Vec<SvgLifecycleIntent>,
    pub(in crate::workbook::edit) svg_preflight:
        HashMap<(usize, usize), super::super::svg_lifecycle::DrawingPreflight>,
    pub(in crate::workbook::edit) svg_projected_owners:
        HashMap<(usize, PictureSelector), super::super::svg_lifecycle::ProjectedSvgOwner>,
    pub(in crate::workbook::edit) svg_staged_payload_bytes: usize,
    /// Materialized part-byte census of `base`, taken when the transaction
    /// stages its first SVG lifecycle intent. The census decodes every part
    /// payload (ADR 0030), so an edit that stages no SVG lifecycle intent never
    /// takes it.
    pub(in crate::workbook::edit) svg_base_part_bytes: Option<usize>,
    #[allow(dead_code, reason = "the facade consumes the native capture budget")]
    capture_budget: Option<CaptureBudget>,
}

#[derive(Debug, Clone)]
pub(super) struct CaptureBudget {
    maximum: usize,
    base_reserved: usize,
    reserved: usize,
    /// Retained bytes for the worksheet-before image and its one changed-part
    /// descriptor. A worksheet is charged once, when the first semantic
    /// action for it is admitted.
    worksheets: BTreeMap<usize, usize>,
    /// Retained bytes for the final semantic cell deltas. The key is the
    /// worksheet position and address, so replacing a payload releases its
    /// old estimate while orthogonal style/structural facets remain charged
    /// alongside the final action.
    cells: BTreeMap<(usize, Address), CellCapture>,
    /// Retained bytes for each newly-created worksheet's final change.
    added_sheets: BTreeMap<usize, usize>,
    /// Retained bytes for the final cell deltas on newly-created worksheets.
    added_cells: BTreeMap<(usize, Address), CellCapture>,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
struct CellCapture {
    payload: usize,
    style_effect: usize,
    font: usize,
    alignment: usize,
    fill: usize,
    border: usize,
    protection: usize,
    quote_prefix: usize,
    number_format: usize,
    structural: usize,
    relationship: usize,
    remove: usize,
    active: bool,
}

/// The typed style facets that still change the immutable source after all
/// pending requests for one address have been composed.  This is deliberately
/// a boolean projection: the capture ledger charges the final retained
/// endpoint, not every intermediate request that an unbounded edit happened
/// to author before a mixed-budget join.
#[derive(Debug, Clone, Copy, Default)]
struct EffectiveTypedFacets {
    font: bool,
    alignment: bool,
    fill: bool,
    border: bool,
    protection: bool,
    quote_prefix: bool,
    number_format: bool,
}

/// The source-equivalence projection produced before a deferred edit is
/// published.  This plan contains only addresses, booleans, and the small
/// action-kind transition.  It is built before any pending map is mutated;
/// applying it then moves superseded values into the undo record instead of
/// cloning payloads, formulas, or style graphs.
#[derive(Debug, Clone, Copy)]
struct NormalizationProjection {
    position: usize,
    address: Address,
    payload_matches: bool,
    style_matches: bool,
    font: bool,
    alignment: bool,
    fill: bool,
    border: bool,
    protection: bool,
    quote_prefix: bool,
    number_format: bool,
    action: ActionNormalization,
}

#[derive(Debug, Clone, Copy)]
enum ActionNormalization {
    Keep,
    Remove,
    RemoveToUpdate,
}

#[derive(Debug)]
struct NormalizationPlan {
    position: usize,
    address: Address,
    action: Option<ActionUndo>,
    font: Option<FontChange>,
    alignment: Option<AlignmentChange>,
    fill: Option<FillChange>,
    border: Option<BorderChange>,
    protection: Option<ProtectionChange>,
    quote_prefix: Option<Option<bool>>,
    number_format: Option<Option<Arc<str>>>,
}

#[derive(Debug)]
enum ActionUndo {
    Removed(Action),
    Update {
        payload: Option<Payload>,
        style: Option<StyleEffect>,
    },
    RemoveToUpdate,
}

#[derive(Debug)]
struct NormalizationUndo {
    plans: Vec<NormalizationPlan>,
    capture_budget: Option<CaptureBudget>,
}

impl NormalizationUndo {
    fn restore(self, edit: &mut Edit) {
        edit.capture_budget = self.capture_budget;
        let mut positions = BTreeSet::new();
        for plan in self.plans.into_iter().rev() {
            positions.insert(plan.position);
            let actions = edit.sheets.entry(plan.position).or_default();
            if let Some(value) = plan.font {
                actions.fonts.insert(plan.address, value);
            }
            if let Some(value) = plan.alignment {
                actions.alignments.insert(plan.address, value);
            }
            if let Some(value) = plan.fill {
                actions.fills.insert(plan.address, value);
            }
            if let Some(value) = plan.border {
                actions.borders.insert(plan.address, value);
            }
            if let Some(value) = plan.protection {
                actions.protections.insert(plan.address, value);
            }
            if let Some(value) = plan.quote_prefix {
                actions.quote_prefixes.insert(plan.address, value);
            }
            if let Some(value) = plan.number_format {
                actions.number_formats.insert(plan.address, value);
            }
            match plan.action {
                Some(ActionUndo::Removed(action)) => {
                    actions.cells.insert(plan.address, action);
                },
                Some(ActionUndo::Update { payload, style }) => {
                    if let Some(Action::Update {
                        payload: current_payload,
                        style: current_style,
                    }) = actions.cells.get_mut(&plan.address)
                    {
                        if let Some(payload) = payload {
                            *current_payload = Some(payload);
                        }
                        if let Some(style) = style {
                            *current_style = Some(style);
                        }
                    }
                },
                Some(ActionUndo::RemoveToUpdate) => {
                    actions.cells.insert(plan.address, Action::Remove);
                },
                None => {},
            }
        }
        for position in positions {
            if let Some(actions) = edit.sheets.get_mut(&position) {
                actions.staged_fill_bytes = actions
                    .fills
                    .values()
                    .map(FillChange::retained_bytes)
                    .fold(0, usize::saturating_add);
                actions.staged_border_bytes = actions
                    .borders
                    .keys()
                    .map(|_| border_staged_entry_bytes())
                    .fold(0, usize::saturating_add);
            }
        }
    }
}

impl CellCapture {
    fn retained_bytes(self) -> usize {
        if !self.active {
            return 0;
        }
        size_of::<Change>()
            .saturating_add(self.payload)
            .saturating_add(self.style_effect)
            .saturating_add(self.font)
            .saturating_add(self.alignment)
            .saturating_add(self.fill)
            .saturating_add(self.border)
            .saturating_add(self.protection)
            .saturating_add(self.quote_prefix)
            .saturating_add(self.number_format)
            .saturating_add(self.structural)
            .saturating_add(self.relationship)
            .saturating_add(self.remove)
    }

    fn apply(&mut self, facet: CaptureFacet, authored_bytes: usize) {
        self.active = true;
        match facet {
            CaptureFacet::Payload => {
                self.payload = authored_bytes;
                self.remove = 0;
            },
            CaptureFacet::Style => {
                self.style_effect = authored_bytes;
                self.font = 0;
                self.alignment = 0;
                self.fill = 0;
                self.border = 0;
                self.protection = 0;
                self.quote_prefix = 0;
                self.number_format = 0;
                self.remove = 0;
            },
            CaptureFacet::Font => {
                self.font = authored_bytes;
                self.remove = 0;
            },
            CaptureFacet::Alignment => {
                self.alignment = authored_bytes;
                self.remove = 0;
            },
            CaptureFacet::Fill => {
                self.fill = authored_bytes;
                self.remove = 0;
            },
            CaptureFacet::Border => {
                self.border = authored_bytes;
                self.remove = 0;
            },
            CaptureFacet::Protection => {
                self.protection = authored_bytes;
                self.remove = 0;
            },
            CaptureFacet::QuotePrefix => {
                self.quote_prefix = authored_bytes;
                self.remove = 0;
            },
            CaptureFacet::NumberFormat => {
                self.number_format = authored_bytes;
                self.remove = 0;
            },
            CaptureFacet::Structural => self.structural = self.structural.max(authored_bytes),
            CaptureFacet::Relationship => {
                self.relationship = self.relationship.saturating_add(authored_bytes);
            },
            CaptureFacet::Remove => {
                self.payload = 0;
                self.style_effect = 0;
                self.font = 0;
                self.alignment = 0;
                self.fill = 0;
                self.border = 0;
                self.protection = 0;
                self.quote_prefix = 0;
                self.number_format = 0;
                self.remove = authored_bytes;
            },
        }
    }

    fn clear(&mut self, facet: CaptureFacet) {
        match facet {
            CaptureFacet::Payload => self.payload = 0,
            CaptureFacet::Style => self.style_effect = 0,
            CaptureFacet::Font => self.font = 0,
            CaptureFacet::Alignment => self.alignment = 0,
            CaptureFacet::Fill => self.fill = 0,
            CaptureFacet::Border => self.border = 0,
            CaptureFacet::Protection => self.protection = 0,
            CaptureFacet::QuotePrefix => self.quote_prefix = 0,
            CaptureFacet::NumberFormat => self.number_format = 0,
            CaptureFacet::Structural => self.structural = 0,
            CaptureFacet::Relationship => self.relationship = 0,
            CaptureFacet::Remove => self.remove = 0,
        }
        self.active = self.payload != 0
            || self.style_effect != 0
            || self.font != 0
            || self.alignment != 0
            || self.fill != 0
            || self.border != 0
            || self.protection != 0
            || self.quote_prefix != 0
            || self.number_format != 0
            || self.structural != 0
            || self.relationship != 0
            || self.remove != 0;
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum CaptureFacet {
    Payload,
    Style,
    Font,
    Alignment,
    Fill,
    Border,
    Protection,
    QuotePrefix,
    NumberFormat,
    Structural,
    Relationship,
    Remove,
}

impl CaptureBudget {
    fn new(base: &Workbook, maximum: usize) -> Self {
        let archive_hint = base.inner.package.exact_source_len().unwrap_or(0);
        let endpoint = workbook_retained_bytes(base, archive_hint);
        // `Patch::retained_bytes` charges the two immutable endpoints and the
        // semantic/native delta buffers. It does not charge the transient
        // `Edit`, `Action`, or `SheetActions` containers, which disappear at
        // commit and used to make this preflight reject valid edits.
        Self {
            maximum,
            base_reserved: endpoint.saturating_mul(2),
            reserved: endpoint.saturating_mul(2),
            worksheets: BTreeMap::new(),
            cells: BTreeMap::new(),
            added_sheets: BTreeMap::new(),
            added_cells: BTreeMap::new(),
        }
    }

    /// Build a one-time ledger for an edit created without a snapshot limit.
    /// Unbounded edits do not pay this census cost while they are being
    /// authored; a mixed-budget join is the first operation that needs their
    /// already-staged native captures admitted against the bounded side.
    fn from_staged(
        base: &Workbook,
        maximum: usize,
        sheets: &BTreeMap<usize, SheetActions>,
        added: &[Added],
    ) -> Result<Self> {
        Self::from_staged_onto(Self::new(base, maximum), base, sheets, added)
    }

    /// [`Self::from_staged`] for a ledger rebuilt over the same `base` as
    /// `previous`: the two immutable endpoint charges are a function of
    /// `base` alone, so they are taken from `previous` instead of walking
    /// the workbook again (record 0377).
    fn rebuilt(
        previous: &Self,
        base: &Workbook,
        sheets: &BTreeMap<usize, SheetActions>,
        added: &[Added],
    ) -> Result<Self> {
        let empty = Self {
            maximum: previous.maximum,
            base_reserved: previous.base_reserved,
            reserved: previous.base_reserved,
            worksheets: BTreeMap::new(),
            cells: BTreeMap::new(),
            added_sheets: BTreeMap::new(),
            added_cells: BTreeMap::new(),
        };
        Self::from_staged_onto(empty, base, sheets, added)
    }

    fn from_staged_onto(
        mut budget: Self,
        base: &Workbook,
        sheets: &BTreeMap<usize, SheetActions>,
        added: &[Added],
    ) -> Result<Self> {
        for (position, actions) in sheets {
            let mut addresses = BTreeSet::new();
            addresses.extend(actions.cells.keys().copied());
            addresses.extend(actions.fonts.keys().copied());
            addresses.extend(actions.alignments.keys().copied());
            addresses.extend(actions.fills.keys().copied());
            addresses.extend(actions.borders.keys().copied());
            addresses.extend(actions.protections.keys().copied());
            addresses.extend(actions.quote_prefixes.keys().copied());
            addresses.extend(actions.number_formats.keys().copied());
            addresses.extend(actions.merges.iter().map(|intent| intent.range().start()));
            addresses.extend(
                actions
                    .hyperlinks
                    .keys()
                    .filter_map(|key| Address::at(key.0, key.1).ok()),
            );
            if addresses.is_empty() {
                continue;
            }
            let source = base
                .sheet(*position)?
                .ok_or_else(|| invalid("worksheet capture target disappeared"))?;
            let store = source.store()?;
            let base_merges = store.merge_ranges().to_vec();
            let mut projected_merges = base_merges.clone();
            for intent in &actions.merges {
                let range = intent.range();
                match intent {
                    MergeIntent::Add(_) => {
                        if !projected_merges.contains(&range) {
                            projected_merges.push(range);
                        }
                    },
                    MergeIntent::Remove(_) => {
                        if let Some(index) = projected_merges.iter().position(|item| *item == range)
                        {
                            projected_merges.remove(index);
                        }
                    },
                }
            }
            let mut effective_merge_starts = BTreeSet::new();
            for range in &base_merges {
                if !projected_merges.contains(range) {
                    effective_merge_starts.insert(range.start());
                }
            }
            for range in &projected_merges {
                if !base_merges.contains(range) {
                    effective_merge_starts.insert(range.start());
                }
            }
            let worksheet =
                worksheet_capture_bytes(base, *position)?.saturating_add(size_of::<PartChange>());
            budget.worksheets.insert(*position, worksheet);
            for address in addresses {
                let action_effective = if let Some(action) = actions.cells.get(&address) {
                    let before = store.entry(address);
                    let before_state = State::read(before, base);
                    let after_state = State::after(before, action, base);
                    before_state != after_state
                        || !action_payload_matches_source(&source, before, address, action)?
                } else {
                    false
                };
                let has_typed = actions.fonts.contains_key(&address)
                    || actions.alignments.contains_key(&address)
                    || actions.fills.contains_key(&address)
                    || actions.borders.contains_key(&address)
                    || actions.protections.contains_key(&address)
                    || actions.quote_prefixes.contains_key(&address)
                    || actions.number_formats.contains_key(&address);
                let effective_format = if !has_typed {
                    None
                } else if let Some(Action::Update {
                    style: Some(style), ..
                }) = actions.cells.get(&address)
                {
                    let raw = match style {
                        StyleEffect::Set(raw) => *raw,
                        StyleEffect::Reset => 0,
                    };
                    Some(base.inner.format(raw)?)
                } else {
                    source
                        .style(address)?
                        .map(|style| style.format())
                        .transpose()?
                };
                let typed = EffectiveTypedFacets {
                    font: actions.fonts.get(&address).is_some_and(|change| {
                        !font_change_matches_source(
                            change,
                            effective_format.as_ref().and_then(Format::font),
                        )
                    }),
                    alignment: actions.alignments.get(&address).is_some_and(|change| {
                        !alignment_change_matches_source(
                            change,
                            effective_format.as_ref().and_then(Format::alignment),
                        )
                    }),
                    fill: actions.fills.get(&address).is_some_and(|change| {
                        !fill_change_matches_source(
                            change,
                            effective_format.as_ref().and_then(Format::fill),
                        )
                    }),
                    border: actions.borders.get(&address).is_some_and(|change| {
                        !border_change_matches_source(
                            change,
                            effective_format.as_ref().and_then(Format::border),
                        )
                    }),
                    protection: actions.protections.get(&address).is_some_and(|change| {
                        !protection_change_matches_source(
                            change,
                            effective_format.as_ref().and_then(Format::protection),
                        )
                    }),
                    quote_prefix: actions.quote_prefixes.get(&address).is_some_and(|value| {
                        !quote_prefix_matches_source(
                            *value,
                            effective_format.as_ref().map(Format::quote_prefix),
                        )
                    }),
                    number_format: actions.number_formats.get(&address).is_some_and(|value| {
                        !number_format_matches_source(
                            value.as_deref(),
                            effective_format
                                .as_ref()
                                .and_then(|format| format.number_format().code()),
                        )
                    }),
                };
                let has_other_facet = typed.font
                    || typed.alignment
                    || typed.fill
                    || typed.border
                    || typed.protection
                    || typed.quote_prefix
                    || typed.number_format
                    || effective_merge_starts.contains(&address)
                    || actions
                        .hyperlinks
                        .keys()
                        .any(|key| Address::at(key.0, key.1).ok() == Some(address));
                if action_effective || has_other_facet {
                    if let Some(capture) = staged_cell_capture(
                        actions,
                        address,
                        action_effective,
                        effective_merge_starts.contains(&address),
                        typed,
                    ) {
                        budget.cells.insert((*position, address), capture);
                    }
                }
            }
            if !budget.cells.keys().any(|key| key.0 == *position) {
                budget.worksheets.remove(position);
            }
        }
        for (index, added) in added.iter().enumerate() {
            let mut addresses = BTreeSet::new();
            addresses.extend(added.actions.cells.keys().copied());
            addresses.extend(added.actions.fonts.keys().copied());
            addresses.extend(added.actions.alignments.keys().copied());
            addresses.extend(added.actions.fills.keys().copied());
            addresses.extend(added.actions.borders.keys().copied());
            addresses.extend(added.actions.protections.keys().copied());
            addresses.extend(added.actions.quote_prefixes.keys().copied());
            addresses.extend(added.actions.number_formats.keys().copied());
            addresses.extend(
                added
                    .actions
                    .merges
                    .iter()
                    .map(|intent| intent.range().start()),
            );
            addresses.extend(
                added
                    .actions
                    .hyperlinks
                    .keys()
                    .filter_map(|key| Address::at(key.0, key.1).ok()),
            );
            budget.added_sheets.insert(
                index,
                size_of::<Change>()
                    .saturating_add(size_of::<PartChange>())
                    .saturating_add(added.name.as_str().len()),
            );
            for address in addresses {
                if let Some(capture) = staged_cell_capture(
                    &added.actions,
                    address,
                    true,
                    true,
                    EffectiveTypedFacets {
                        font: true,
                        alignment: true,
                        fill: true,
                        border: true,
                        protection: true,
                        quote_prefix: true,
                        number_format: true,
                    },
                ) {
                    budget.added_cells.insert((index, address), capture);
                }
            }
        }
        budget.recompute_reserved();
        Ok(budget)
    }

    fn replace(&mut self, previous: usize, next: usize) -> Result<()> {
        let observed = self.reserved.saturating_sub(previous).saturating_add(next);
        if observed > self.maximum {
            return Err(Error::Limit {
                kind: super::super::LimitKind::SnapshotBytes.into(),
                observed: u64::try_from(observed).unwrap_or(u64::MAX),
                maximum: u64::try_from(self.maximum).unwrap_or(u64::MAX),
            });
        }
        self.reserved = observed;
        Ok(())
    }

    fn recompute_reserved(&mut self) {
        let cells = self
            .cells
            .values()
            .map(|capture| capture.retained_bytes())
            .fold(0usize, usize::saturating_add);
        let added_cells = self
            .added_cells
            .values()
            .map(|capture| capture.retained_bytes())
            .fold(0usize, usize::saturating_add);
        let worksheets = self
            .worksheets
            .values()
            .copied()
            .fold(0usize, usize::saturating_add);
        let added_sheets = self
            .added_sheets
            .values()
            .copied()
            .fold(0usize, usize::saturating_add);
        self.reserved = self
            .base_reserved
            .saturating_add(worksheets)
            .saturating_add(cells)
            .saturating_add(added_sheets)
            .saturating_add(added_cells);
    }

    fn joined(
        &self,
        other: &Self,
        added_offset: usize,
    ) -> std::result::Result<Self, (usize, usize)> {
        let mut joined = self.clone();
        joined.maximum = self.maximum.min(other.maximum);
        for (position, bytes) in &other.worksheets {
            joined
                .worksheets
                .entry(*position)
                .and_modify(|current| *current = (*current).max(*bytes))
                .or_insert(*bytes);
        }
        for (key, capture) in &other.cells {
            joined
                .cells
                .entry(*key)
                .and_modify(|current| {
                    current.payload = current.payload.max(capture.payload);
                    current.style_effect = current.style_effect.max(capture.style_effect);
                    current.font = current.font.max(capture.font);
                    current.alignment = current.alignment.max(capture.alignment);
                    current.fill = current.fill.max(capture.fill);
                    current.border = current.border.max(capture.border);
                    current.protection = current.protection.max(capture.protection);
                    current.quote_prefix = current.quote_prefix.max(capture.quote_prefix);
                    current.number_format = current.number_format.max(capture.number_format);
                    current.structural = current.structural.max(capture.structural);
                    current.relationship =
                        current.relationship.saturating_add(capture.relationship);
                    current.remove = current.remove.max(capture.remove);
                    current.active |= capture.active;
                })
                .or_insert(*capture);
        }
        for (index, bytes) in &other.added_sheets {
            let index = (*index).saturating_add(added_offset);
            joined
                .added_sheets
                .entry(index)
                .and_modify(|current| *current = (*current).max(*bytes))
                .or_insert(*bytes);
        }
        for (key, capture) in &other.added_cells {
            let key = (key.0.saturating_add(added_offset), key.1);
            joined
                .added_cells
                .entry(key)
                .and_modify(|current| {
                    current.payload = current.payload.max(capture.payload);
                    current.style_effect = current.style_effect.max(capture.style_effect);
                    current.font = current.font.max(capture.font);
                    current.alignment = current.alignment.max(capture.alignment);
                    current.fill = current.fill.max(capture.fill);
                    current.border = current.border.max(capture.border);
                    current.protection = current.protection.max(capture.protection);
                    current.quote_prefix = current.quote_prefix.max(capture.quote_prefix);
                    current.number_format = current.number_format.max(capture.number_format);
                    current.structural = current.structural.max(capture.structural);
                    current.relationship =
                        current.relationship.saturating_add(capture.relationship);
                    current.remove = current.remove.max(capture.remove);
                    current.active |= capture.active;
                })
                .or_insert(*capture);
        }
        joined.recompute_reserved();
        if joined.reserved > joined.maximum {
            return Err((joined.reserved, joined.maximum));
        }
        Ok(joined)
    }

    fn check_patch(&self, patch: &Patch) -> Result<()> {
        let observed = patch.retained_bytes();
        if observed > self.maximum {
            return Err(Error::Limit {
                kind: super::super::LimitKind::SnapshotBytes.into(),
                observed: u64::try_from(observed).unwrap_or(u64::MAX),
                maximum: u64::try_from(self.maximum).unwrap_or(u64::MAX),
            });
        }
        Ok(())
    }

    fn admit_cell(
        &mut self,
        position: usize,
        address: Address,
        worksheet_bytes: usize,
        facet: CaptureFacet,
        authored_bytes: usize,
    ) -> Result<CaptureReservation> {
        let key = (position, address);
        let previous = self.cells.get(&key).copied();
        let worksheet_previous = self.worksheets.get(&position).copied();
        let worksheet_next = if worksheet_previous.is_none() {
            worksheet_bytes.saturating_add(size_of::<PartChange>())
        } else {
            worksheet_previous.unwrap_or(0)
        };
        let mut next = previous.unwrap_or_default();
        next.apply(facet, authored_bytes);
        let old_total = previous
            .map_or(0, CellCapture::retained_bytes)
            .saturating_add(worksheet_previous.unwrap_or(0));
        let new_total = next.retained_bytes().saturating_add(worksheet_next);
        self.replace(old_total, new_total)?;
        self.cells.insert(key, next);
        if worksheet_previous.is_none() {
            // Keep the descriptor in the map exactly as it was admitted.
            // Rollback and recomputation must charge the same worksheet
            // owner; storing only `worksheet_bytes` made the first-cell
            // reservation impossible to roll back and dropped the
            // descriptor from joined ledgers.
            self.worksheets.insert(position, worksheet_next);
        }
        Ok(CaptureReservation {
            key,
            previous,
            worksheet_previous,
            next,
            worksheet_next,
            active: true,
        })
    }

    fn admit_cell_relationship(
        &mut self,
        position: usize,
        address: Address,
        worksheet_bytes: usize,
        previous_relationship: usize,
        authored_bytes: usize,
    ) -> Result<CaptureReservation> {
        let key = (position, address);
        let previous = self.cells.get(&key).copied();
        let worksheet_previous = self.worksheets.get(&position).copied();
        let worksheet_next = if worksheet_previous.is_none() {
            worksheet_bytes.saturating_add(size_of::<PartChange>())
        } else {
            worksheet_previous.unwrap_or(0)
        };
        let mut next = previous.unwrap_or_default();
        if previous_relationship > next.relationship {
            return Err(invalid("worksheet hyperlink capture ledger underflow"));
        }
        next.relationship = next
            .relationship
            .saturating_sub(previous_relationship)
            .saturating_add(authored_bytes);
        next.active = true;
        let old_total = previous
            .map_or(0, CellCapture::retained_bytes)
            .saturating_add(worksheet_previous.unwrap_or(0));
        let new_total = next.retained_bytes().saturating_add(worksheet_next);
        self.replace(old_total, new_total)?;
        self.cells.insert(key, next);
        if worksheet_previous.is_none() {
            self.worksheets.insert(position, worksheet_next);
        }
        Ok(CaptureReservation {
            key,
            previous,
            worksheet_previous,
            next,
            worksheet_next,
            active: true,
        })
    }

    pub(super) fn admit_new_sheet(&mut self, index: usize, authored_bytes: usize) -> Result<()> {
        let previous = self.added_sheets.get(&index).copied().unwrap_or(0);
        let next = size_of::<Change>()
            .saturating_add(size_of::<PartChange>())
            .saturating_add(authored_bytes);
        self.replace(previous, next)?;
        self.added_sheets.insert(index, next);
        Ok(())
    }

    pub(super) fn charge_new_sheet(&mut self, index: usize, authored_bytes: usize) -> Result<()> {
        self.admit_new_sheet(index, authored_bytes)
    }

    pub(super) fn charge_new_sheet_cell(
        &mut self,
        index: usize,
        address: Address,
        facet: CaptureFacet,
        authored_bytes: usize,
    ) -> Result<AddedCaptureReservation> {
        let key = (index, address);
        let previous = self.added_cells.get(&key).copied();
        let mut next = previous.unwrap_or_default();
        next.apply(facet, authored_bytes);
        self.replace(
            previous.map_or(0, CellCapture::retained_bytes),
            next.retained_bytes(),
        )?;
        self.added_cells.insert(key, next);
        Ok(AddedCaptureReservation {
            key,
            previous,
            next,
            active: true,
        })
    }

    pub(super) fn rollback_added_cell(&mut self, mut reservation: AddedCaptureReservation) {
        if !reservation.active {
            return;
        }
        if self.added_cells.get(&reservation.key).copied() != Some(reservation.next) {
            return;
        }
        let current = reservation.next.retained_bytes();
        let previous = reservation.previous.map_or(0, CellCapture::retained_bytes);
        self.reserved = self
            .reserved
            .saturating_sub(current)
            .saturating_add(previous);
        if let Some(previous) = reservation.previous.take() {
            self.added_cells.insert(reservation.key, previous);
        } else {
            self.added_cells.remove(&reservation.key);
        }
    }

    pub(super) fn release_added_cell(&mut self, index: usize, address: Address) {
        self.added_cells.remove(&(index, address));
        self.recompute_reserved();
    }
}

/// A successful cell admission can be rolled back when a source validation
/// that necessarily follows the admission proves the operation is not a
/// mutation (for example, `set_cached` on a non-formula cell).
#[derive(Debug)]
pub(crate) struct CaptureReservation {
    key: (usize, Address),
    previous: Option<CellCapture>,
    worksheet_previous: Option<usize>,
    next: CellCapture,
    worksheet_next: usize,
    active: bool,
}

#[derive(Debug)]
pub(crate) struct AddedCaptureReservation {
    key: (usize, Address),
    previous: Option<CellCapture>,
    next: CellCapture,
    active: bool,
}

impl AddedCaptureReservation {
    pub(super) fn inactive() -> Self {
        Self {
            key: (0, Address::at(0, 0).expect("zero cell address")),
            previous: None,
            next: CellCapture {
                payload: 0,
                style_effect: 0,
                font: 0,
                alignment: 0,
                fill: 0,
                border: 0,
                protection: 0,
                quote_prefix: 0,
                number_format: 0,
                structural: 0,
                relationship: 0,
                remove: 0,
                active: false,
            },
            active: false,
        }
    }
}

impl Edit {
    pub(super) const fn capture_budgeted(&self) -> bool {
        self.capture_budget.is_some()
    }

    pub(crate) fn new(base: Workbook) -> Result<Self> {
        Self::new_with_options(base, super::super::EditOptions::default())
    }

    pub(crate) fn new_with_options(
        base: Workbook,
        options: super::super::EditOptions,
    ) -> Result<Self> {
        let capture_budget = options
            .max_retained_bytes_value()
            .map(|maximum| CaptureBudget::new(&base, maximum));
        if let Some(budget) = capture_budget.as_ref() {
            if budget.reserved > budget.maximum {
                return Err(Error::Limit {
                    kind: super::super::LimitKind::SnapshotBytes.into(),
                    observed: u64::try_from(budget.reserved).unwrap_or(u64::MAX),
                    maximum: u64::try_from(budget.maximum).unwrap_or(u64::MAX),
                });
            }
        }
        Ok(Self {
            base,
            panes: None,
            defined_names: None,
            drawings: BTreeMap::new(),
            active: None,
            order: None,
            sheets: BTreeMap::new(),
            added: Vec::new(),
            removed: BTreeSet::new(),
            cross_workbook_scalar: false,
            svg_lifecycle: Vec::new(),
            svg_preflight: HashMap::new(),
            svg_projected_owners: HashMap::new(),
            svg_staged_payload_bytes: 0,
            svg_base_part_bytes: None,
            capture_budget,
        })
    }

    /// Reserve the native capture needed by one worksheet operation before
    /// its action map, changed store, or compaction output can allocate.
    #[allow(dead_code, reason = "called by the XLSX facade coalescer")]
    pub(crate) fn preflight_cell_capture(
        &mut self,
        position: usize,
        address: Address,
        authored_bytes: usize,
    ) -> Result<CaptureReservation> {
        self.preflight_cell_capture_with_facet(
            position,
            address,
            CaptureFacet::Payload,
            authored_bytes,
        )
    }

    pub(crate) fn preflight_cell_capture_with_facet(
        &mut self,
        position: usize,
        address: Address,
        facet: CaptureFacet,
        authored_bytes: usize,
    ) -> Result<CaptureReservation> {
        if self.capture_budget.is_none() {
            return Ok(CaptureReservation {
                key: (position, address),
                previous: None,
                worksheet_previous: None,
                next: CellCapture::default(),
                worksheet_next: 0,
                active: false,
            });
        }
        let sheet = self
            .base
            .inner
            .sheets
            .get(position)
            .ok_or_else(|| invalid("worksheet capture target disappeared"))?;
        let worksheet_bytes = if !self
            .capture_budget
            .as_ref()
            .is_some_and(|budget| budget.worksheets.contains_key(&position))
        {
            match self.base.inner.package.part_metadata(&sheet.part_uri) {
                Some(metadata) if metadata.payload_is_decoded() => self
                    .base
                    .inner
                    .package
                    .get_part(&sheet.part_uri)?
                    .blob()
                    .len(),
                // A deferred source payload must not be decoded merely to
                // decide whether the capture budget admits the edit. Use its
                // central-directory uncompressed declaration as the bound;
                // the later payload force still validates the actual bytes.
                Some(metadata) => metadata.payload_len()?,
                None => return Err(invalid("worksheet capture part disappeared")),
            }
        } else {
            0
        };
        self.capture_budget
            .as_mut()
            .expect("capture budget checked above")
            .admit_cell(position, address, worksheet_bytes, facet, authored_bytes)
    }

    /// Admit a hyperlink replacement against the final relationship set for
    /// this cell. More than one range can start at the same cell, so a new
    /// relationship is additive while replacing one staged key releases only
    /// that key's prior retained payload.
    fn preflight_hyperlink_capture(
        &mut self,
        position: usize,
        address: Address,
        previous_bytes: usize,
        authored_bytes: usize,
    ) -> Result<CaptureReservation> {
        if self.capture_budget.is_none() {
            return Ok(CaptureReservation {
                key: (position, address),
                previous: None,
                worksheet_previous: None,
                next: CellCapture::default(),
                worksheet_next: 0,
                active: false,
            });
        }
        let sheet = self
            .base
            .inner
            .sheets
            .get(position)
            .ok_or_else(|| invalid("worksheet capture target disappeared"))?;
        let worksheet_bytes = if !self
            .capture_budget
            .as_ref()
            .is_some_and(|budget| budget.worksheets.contains_key(&position))
        {
            match self.base.inner.package.part_metadata(&sheet.part_uri) {
                Some(metadata) => metadata.payload_len()?,
                None => return Err(invalid("worksheet capture part disappeared")),
            }
        } else {
            0
        };
        self.capture_budget
            .as_mut()
            .expect("capture budget checked above")
            .admit_cell_relationship(
                position,
                address,
                worksheet_bytes,
                previous_bytes,
                authored_bytes,
            )
    }

    /// Admit a payload captured by a transfer planner before its action map is
    /// populated.  Transfer planners build their complete action list first,
    /// so a failed admission can release every earlier reservation atomically.
    pub(crate) fn preflight_transfer_payload(
        &mut self,
        position: usize,
        address: Address,
        authored_bytes: usize,
    ) -> Result<CaptureReservation> {
        self.preflight_cell_capture_with_facet(
            position,
            address,
            CaptureFacet::Payload,
            authored_bytes,
        )
    }

    /// Admit the semantic facets represented by one transfer action.  The
    /// action is still only a plan at this point; callers insert it after all
    /// actions in the transfer have passed admission.
    pub(crate) fn preflight_transfer_action(
        &mut self,
        position: usize,
        address: Address,
        action: &Action,
    ) -> Result<Vec<CaptureReservation>> {
        let mut reservations = Vec::with_capacity(2);
        let result = (|| -> Result<()> {
            match action {
                Action::Remove => reservations.push(self.preflight_cell_capture_with_facet(
                    position,
                    address,
                    CaptureFacet::Remove,
                    size_of::<Action>(),
                )?),
                Action::Update { payload, style } => {
                    if let Some(payload) = payload {
                        reservations.push(self.preflight_cell_capture_with_facet(
                            position,
                            address,
                            CaptureFacet::Payload,
                            payload_capture_bytes(payload),
                        )?);
                    }
                    if style.is_some() {
                        reservations.push(self.preflight_cell_capture_with_facet(
                            position,
                            address,
                            CaptureFacet::Style,
                            size_of::<StyleEffect>(),
                        )?);
                    }
                },
            }
            Ok(())
        })();
        if let Err(error) = result {
            self.rollback_capture_reservations(&mut reservations);
            return Err(error);
        }
        Ok(reservations)
    }

    /// Release reservations accumulated by an all-or-nothing transfer plan.
    pub(crate) fn rollback_capture_reservations(
        &mut self,
        reservations: &mut Vec<CaptureReservation>,
    ) {
        while let Some(reservation) = reservations.pop() {
            self.rollback_cell_capture(reservation);
        }
    }

    /// Release one cell facet when a pending action is cancelled before it is
    /// published (for example, a hyperlink put followed by its removal).
    pub(crate) fn release_cell_capture_facet(
        &mut self,
        position: usize,
        address: Address,
        facet: CaptureFacet,
    ) {
        let Some(budget) = self.capture_budget.as_mut() else {
            return;
        };
        let key = (position, address);
        let Some(capture) = budget.cells.get_mut(&key) else {
            return;
        };
        capture.clear(facet);
        let empty = !capture.active;
        if empty {
            budget.cells.remove(&key);
        }
        if !budget.cells.keys().any(|candidate| candidate.0 == position) {
            budget.worksheets.remove(&position);
        }
        budget.recompute_reserved();
    }

    /// Reconcile the relationship facet with the final pending hyperlink
    /// actions for one start cell. Range keys are independent even when they
    /// share that start, so clearing one key must not clear its siblings.
    fn recompute_relationship_capture(&mut self, position: usize, address: Address) {
        let Some(budget) = self.capture_budget.as_mut() else {
            return;
        };
        let retained = self.sheets.get(&position).map_or(0, |actions| {
            actions
                .hyperlinks
                .iter()
                .filter(|((row, column, _, _), _)| Address::at(*row, *column).ok() == Some(address))
                .map(|(_, action)| hyperlink_action_retained_bytes(action))
                .fold(0usize, usize::saturating_add)
        });
        let key = (position, address);
        let Some(capture) = budget.cells.get_mut(&key) else {
            return;
        };
        capture.relationship = retained;
        capture.active = capture.payload != 0
            || capture.style_effect != 0
            || capture.font != 0
            || capture.alignment != 0
            || capture.fill != 0
            || capture.border != 0
            || capture.protection != 0
            || capture.quote_prefix != 0
            || capture.number_format != 0
            || capture.structural != 0
            || capture.relationship != 0
            || capture.remove != 0;
        if !capture.active {
            budget.cells.remove(&key);
        }
        if !budget.cells.keys().any(|candidate| candidate.0 == position) {
            budget.worksheets.remove(&position);
        }
        budget.recompute_reserved();
    }

    pub(crate) fn release_cell_capture_all(&mut self, position: usize, address: Address) {
        let Some(budget) = self.capture_budget.as_mut() else {
            return;
        };
        budget.cells.remove(&(position, address));
        if !budget.cells.keys().any(|key| key.0 == position) {
            budget.worksheets.remove(&position);
        }
        budget.recompute_reserved();
    }

    fn recompute_capture_budget(&mut self) {
        if let Some(budget) = self.capture_budget.as_mut() {
            budget.recompute_reserved();
        }
    }

    /// Drop the cell facets superseded by a complete transfer replacement.
    /// Relationship and structural captures belong to independent worksheet
    /// graphs and therefore remain charged. Recompute once for the whole
    /// operation rather than clearing each facet through the O(n) public
    /// helper repeatedly.
    fn release_transfer_capture(&mut self, position: usize, address: Address) {
        let Some(budget) = self.capture_budget.as_mut() else {
            return;
        };
        let Some(capture) = budget.cells.get_mut(&(position, address)) else {
            return;
        };
        capture.payload = 0;
        capture.style_effect = 0;
        capture.font = 0;
        capture.alignment = 0;
        capture.fill = 0;
        capture.border = 0;
        capture.protection = 0;
        capture.quote_prefix = 0;
        capture.number_format = 0;
        capture.remove = 0;
        capture.active = capture.structural != 0 || capture.relationship != 0;
        if !capture.active {
            budget.cells.remove(&(position, address));
        }
        if !budget.cells.keys().any(|key| key.0 == position) {
            budget.worksheets.remove(&position);
        }
    }

    /// Admit a cell's native before-image before a caller reads the projected
    /// state that may decode a deferred worksheet. The semantic worksheet
    /// methods perform the precise replacement charge again when they stage
    /// the final action.
    #[doc(hidden)]
    pub fn reserve_cell_capture(
        &mut self,
        position: usize,
        address: Address,
        authored_bytes: usize,
    ) -> Result<()> {
        self.preflight_cell_capture(position, address, authored_bytes)
            .map(drop)
    }

    pub(crate) fn rollback_cell_capture(&mut self, mut reservation: CaptureReservation) {
        if !reservation.active {
            return;
        }
        let Some(budget) = &mut self.capture_budget else {
            return;
        };
        if budget.cells.get(&reservation.key).copied() != Some(reservation.next)
            || budget.worksheets.get(&reservation.key.0).copied()
                != Some(reservation.worksheet_next)
        {
            return;
        }
        let current = reservation
            .next
            .retained_bytes()
            .saturating_add(reservation.worksheet_next);
        let previous = reservation
            .previous
            .map_or(0, CellCapture::retained_bytes)
            .saturating_add(reservation.worksheet_previous.unwrap_or(0));
        budget.reserved = budget
            .reserved
            .saturating_sub(current)
            .saturating_add(previous);
        if let Some(previous) = reservation.previous.take() {
            budget.cells.insert(reservation.key, previous);
        } else {
            budget.cells.remove(&reservation.key);
        }
        if let Some(previous) = reservation.worksheet_previous.take() {
            budget.worksheets.insert(reservation.key.0, previous);
        } else {
            budget.worksheets.remove(&reservation.key.0);
        }
    }

    /// Take, once, the base census the SVG lifecycle planner checks at commit.
    fn take_svg_base_census(&mut self) -> Result<()> {
        if self.svg_base_part_bytes.is_none() {
            self.svg_base_part_bytes = Some(svg_plan::materialized_part_bytes(&self.base)?);
        }
        Ok(())
    }

    /// Stage one fill after checking the aggregate retained gradient payload.
    ///
    /// The native stylesheet rewrite has its own output and staged-resource
    /// guards. This admission closes the earlier semantic gap where many
    /// cloned `Arc`-backed gradient requests could accumulate in a transaction
    /// before lowering reached that codec. The incoming request is checked
    /// before it is cloned; the merged request is checked again after a
    /// replacement or property merge has settled its retained payload.
    pub(super) fn stage_fill(
        &mut self,
        position: usize,
        address: Address,
        change: &FillChange,
    ) -> Result<()> {
        let maximum = style_staged_memory_limit(&self.base);
        let incoming = change.retained_bytes();
        enforce_style_staged_memory(incoming, maximum, "worksheet fill staged bytes")?;

        let previous_bytes = self
            .sheets
            .get(&position)
            .and_then(|actions| actions.fills.get(&address))
            .map(FillChange::retained_bytes)
            .unwrap_or(0);
        let previous = self
            .sheets
            .get(&position)
            .and_then(|actions| actions.fills.get(&address))
            .cloned();
        let mut merged = previous.unwrap_or_else(FillChange::new);
        merged.merge(change);

        let merged_bytes = merged.retained_bytes();
        let current_sheet_bytes = self
            .sheets
            .get(&position)
            .map_or(0, |actions| actions.staged_fill_bytes);
        let updated_sheet_bytes = current_sheet_bytes
            .checked_sub(previous_bytes)
            .and_then(|bytes| bytes.checked_add(merged_bytes))
            .ok_or_else(|| invalid("worksheet fill retained bytes overflow"))?;
        let current_total = staged_style_bytes(&self.sheets)?;
        let retained = current_total
            .checked_sub(previous_bytes)
            .and_then(|bytes| bytes.checked_add(merged_bytes))
            .ok_or_else(|| invalid("worksheet fill retained bytes overflow"))?;
        enforce_style_staged_memory(retained, maximum, "worksheet style staged bytes")?;

        let actions = self.sheets.entry(position).or_default();
        actions.fills.insert(address, merged);
        actions.staged_fill_bytes = updated_sheet_bytes;
        Ok(())
    }

    /// Stage one border after admitting the fixed border value, map node, and
    /// lowering-copy charge against the aggregate deferred-style budget.
    pub(super) fn stage_border(
        &mut self,
        position: usize,
        address: Address,
        change: &BorderChange,
    ) -> Result<()> {
        let maximum = style_staged_memory_limit(&self.base);
        let is_new = self
            .sheets
            .get(&position)
            .is_none_or(|actions| !actions.borders.contains_key(&address));
        if is_new {
            let retained = staged_style_bytes(&self.sheets)?
                .checked_add(border_staged_entry_bytes())
                .ok_or_else(|| invalid("worksheet border retained bytes overflow"))?;
            enforce_style_staged_memory(retained, maximum, "worksheet style staged bytes")?;
        }
        let added_border_bytes = if is_new {
            border_staged_entry_bytes()
        } else {
            0
        };
        let updated_border_bytes = self
            .sheets
            .get(&position)
            .map_or(0, |actions| actions.staged_border_bytes)
            .checked_add(added_border_bytes)
            .ok_or_else(|| invalid("worksheet border retained bytes overflow"))?;
        let previous = self
            .sheets
            .get(&position)
            .and_then(|actions| actions.borders.get(&address))
            .cloned();
        let mut merged = previous.unwrap_or_else(BorderChange::new);
        merged.merge(change);
        let actions = self.sheets.entry(position).or_default();
        actions.borders.insert(address, merged);
        actions.staged_border_bytes = updated_border_bytes;
        Ok(())
    }

    /// Discard one deferred fill before a complete style, reset, or removal
    /// action supersedes that cell. The counter is adjusted before the map is
    /// mutated so a failed invariant check cannot partially change the edit.
    pub(super) fn discard_fill(&mut self, position: usize, address: Address) -> Result<()> {
        let Some(previous_bytes) = self
            .sheets
            .get(&position)
            .and_then(|actions| actions.fills.get(&address))
            .map(FillChange::retained_bytes)
        else {
            return Ok(());
        };
        let actions = self
            .sheets
            .get_mut(&position)
            .ok_or_else(|| invalid("worksheet fill staging disappeared"))?;
        let updated = actions
            .staged_fill_bytes
            .checked_sub(previous_bytes)
            .ok_or_else(|| invalid("worksheet fill retained bytes underflow"))?;
        actions.fills.remove(&address);
        actions.staged_fill_bytes = updated;
        Ok(())
    }

    /// Discard one deferred border before a complete style, reset, or removal
    /// action supersedes that cell.
    pub(super) fn discard_border(&mut self, position: usize, address: Address) -> Result<()> {
        if !self
            .sheets
            .get(&position)
            .is_some_and(|actions| actions.borders.contains_key(&address))
        {
            return Ok(());
        }
        let actions = self
            .sheets
            .get_mut(&position)
            .ok_or_else(|| invalid("worksheet border staging disappeared"))?;
        actions.staged_border_bytes = actions
            .staged_border_bytes
            .checked_sub(border_staged_entry_bytes())
            .ok_or_else(|| invalid("worksheet border retained bytes underflow"))?;
        actions.borders.remove(&address);
        Ok(())
    }

    pub(super) fn stage_svg_attach(
        &mut self,
        position: usize,
        selector: PictureSelector,
        input: SvgInput<'_>,
    ) -> Result<()> {
        self.stage_svg_attach_with_admission(position, selector, input, || Ok(()))
    }

    #[cfg(test)]
    pub(in crate::workbook::edit) fn stage_svg_attach_after_preflight_failure_for_test(
        &mut self,
        position: usize,
        selector: PictureSelector,
        input: SvgInput<'_>,
    ) -> Result<()> {
        self.stage_svg_attach_with_admission(position, selector, input, || {
            Err(invalid("deterministic SVG staging admission failure"))
        })
    }

    fn stage_svg_attach_with_admission<F>(
        &mut self,
        position: usize,
        selector: PictureSelector,
        input: SvgInput<'_>,
        admit: F,
    ) -> Result<()>
    where
        F: FnOnce() -> Result<()>,
    {
        guard::no_removal(self, "SVG lifecycle")?;
        if self.svg_lifecycle.len() >= svg_plan::MAX_SVG_LIFECYCLE_INTENTS {
            return Err(invalid(format!(
                "SVG lifecycle intents exceed {}",
                svg_plan::MAX_SVG_LIFECYCLE_INTENTS
            )));
        }
        let bytes = input.as_bytes();
        if bytes.is_empty() {
            return Err(invalid("SVG payload cannot be empty"));
        }
        let input_limit = svg_plan::payload_limit(&self.base).min(MAX_SVG_INPUT_BYTES);
        if bytes.len() > input_limit {
            return Err(Error::Limit {
                kind: super::super::LimitKind::SvgInputBytes.into(),
                observed: u64::try_from(bytes.len()).unwrap_or(u64::MAX),
                maximum: u64::try_from(input_limit).unwrap_or(u64::MAX),
            });
        }
        self.take_svg_base_census()?;
        svg_plan::check_payload_budget(&self.base, self.svg_staged_payload_bytes, bytes.len())?;
        let prepared = svg_plan::preflight_with_pending(
            &self.base,
            position,
            selector,
            true,
            &mut self.svg_preflight,
            &mut self.svg_projected_owners,
        )?;
        self.svg_lifecycle
            .try_reserve(1)
            .map_err(|source| allocation("SVG lifecycle intents", source))?;
        let mut owned = Vec::new();
        owned
            .try_reserve_exact(bytes.len())
            .map_err(|source| allocation("SVG attachment input", source))?;
        admit()?;
        owned.extend_from_slice(bytes);
        let staged_payload_bytes = self
            .svg_staged_payload_bytes
            .checked_add(bytes.len())
            .ok_or_else(|| invalid("SVG lifecycle staged payload bytes overflow"))?;
        let intent = SvgLifecycleIntent {
            position,
            selector,
            attach: true,
            payload: Some(Arc::new(owned)),
        };
        prepared.commit(&mut self.svg_preflight, &mut self.svg_projected_owners);
        self.svg_lifecycle.push(intent);
        self.svg_staged_payload_bytes = staged_payload_bytes;
        Ok(())
    }

    pub(super) fn stage_svg_detach(
        &mut self,
        position: usize,
        selector: PictureSelector,
    ) -> Result<bool> {
        guard::no_removal(self, "SVG lifecycle")?;
        if self.svg_lifecycle.len() >= svg_plan::MAX_SVG_LIFECYCLE_INTENTS {
            return Err(invalid(format!(
                "SVG lifecycle intents exceed {}",
                svg_plan::MAX_SVG_LIFECYCLE_INTENTS
            )));
        }
        let prepared = svg_plan::preflight_with_pending(
            &self.base,
            position,
            selector,
            false,
            &mut self.svg_preflight,
            &mut self.svg_projected_owners,
        )?;
        if !prepared.is_effective() {
            prepared.commit(&mut self.svg_preflight, &mut self.svg_projected_owners);
            return Ok(false);
        }
        self.take_svg_base_census()?;
        self.svg_lifecycle
            .try_reserve(1)
            .map_err(|source| allocation("SVG lifecycle intents", source))?;
        let intent = SvgLifecycleIntent {
            position,
            selector,
            attach: false,
            payload: None,
        };
        prepared.commit(&mut self.svg_preflight, &mut self.svg_projected_owners);
        self.svg_lifecycle.push(intent);
        Ok(true)
    }

    /// Create or replace the complete persisted Office Add-in task-pane graph.
    ///
    /// The owned model is moved into this isolated transaction. Its worksheet
    /// `appRef` dependencies are checked against the transaction's combined
    /// final state before any package part is changed.
    pub fn put_task_panes(
        &mut self,
        panes: common_web::Panes,
        conformance: common_web::Conformance,
    ) -> Result<&mut Self> {
        guard::no_removal(self, "task panes")?;
        self.panes = Some(PanesAction::Put { panes, conformance });
        Ok(self)
    }

    /// Remove persisted task panes when no effective worksheet binding would
    /// dangle in the same transaction.
    pub fn remove_task_panes(&mut self) -> Result<&mut Self> {
        guard::no_removal(self, "task panes")?;
        self.panes = Some(PanesAction::Remove);
        Ok(self)
    }

    /// Replace the complete inert workbook defined-name catalog.
    ///
    /// The staged catalog is validated against the current sheet count and
    /// rewritten without rebuilding unrelated workbook XML. Structural tab
    /// edits are deliberately kept separate so local scopes cannot silently
    /// change meaning.
    pub fn replace_defined_names(&mut self, names: Vec<raw::DefinedName>) -> Result<&mut Self> {
        guard::no_removal(self, "defined names")?;
        let workbook = self
            .base
            .inner
            .package
            .get_part(&self.base.inner.workbook_uri)?;
        let candidate = raw::catalog_edit::replace_defined_names(workbook.blob(), &names)?;
        let parsed = raw::parse_catalog(&candidate, &self.base.inner.package.read_limits())?;
        if parsed.defined_names != names {
            return Err(invalid("defined-name authoring verification failed"));
        }
        self.defined_names = (names.as_slice() != self.base.inner.defined_names.as_ref())
            .then(|| names.into_boxed_slice());
        Ok(self)
    }

    /// Remove every direct workbook defined name.
    pub fn clear_defined_names(&mut self) -> Result<&mut Self> {
        self.replace_defined_names(Vec::new())
    }

    /// Append a validated worksheet and borrow its transaction-local editor.
    ///
    /// The returned handle can populate cells and properties before the one
    /// atomic commit. Native sheet IDs, relationship IDs, and part names are
    /// allocated deterministically at commit and never enter the public API.
    pub fn add<T>(&mut self, name: T) -> Result<NewSheet<'_>>
    where
        T: TryInto<Name>,
        Error: From<T::Error>,
    {
        self.add_placed(name, Placement::Tail)
    }

    /// Insert a validated worksheet immediately before a semantic anchor.
    ///
    /// The anchor accepts the ordinary name or checked zero-based selector.
    /// `Ok(None)` means it did not resolve in the source snapshot.
    pub fn add_before<'e, 's, T>(
        &'e mut self,
        name: T,
        anchor: impl Into<Selector<'s>>,
    ) -> Result<Option<NewSheet<'e>>>
    where
        T: TryInto<Name>,
        Error: From<T::Error>,
    {
        guard::no_removal(self, "transaction")?;
        let Some(anchor) = Snapshot::new(&self.base).tab(anchor)? else {
            return Ok(None);
        };
        self.add_placed(name, Placement::Before(anchor.position()))
            .map(Some)
    }

    /// Insert a validated worksheet immediately after a semantic anchor.
    ///
    /// Multiple additions at one anchor retain call order. `Ok(None)` means
    /// the name or checked zero-based anchor did not resolve.
    pub fn add_after<'e, 's, T>(
        &'e mut self,
        name: T,
        anchor: impl Into<Selector<'s>>,
    ) -> Result<Option<NewSheet<'e>>>
    where
        T: TryInto<Name>,
        Error: From<T::Error>,
    {
        guard::no_removal(self, "transaction")?;
        let Some(anchor) = Snapshot::new(&self.base).tab(anchor)? else {
            return Ok(None);
        };
        self.add_placed(name, Placement::After(anchor.position()))
            .map(Some)
    }

    /// Select a worksheet for short transaction-scoped operations.
    pub fn sheet<'e, 's>(
        &'e mut self,
        selector: impl Into<Selector<'s>>,
    ) -> Result<Option<Sheet<'e>>> {
        guard::no_removal(self, "transaction")?;
        let sheet = Snapshot::new(&self.base).worksheet(selector)?;
        let Some(sheet) = sheet else {
            return Ok(None);
        };
        let position = sheet.position();
        Ok(Some(Sheet {
            edit: self,
            position,
        }))
    }

    /// Select any workbook sheet tab by its developer-facing name or checked
    /// zero-based position. Unlike [`Self::sheet`], this entry point also
    /// accepts chart, dialog, and macro sheets because visibility belongs to
    /// the workbook catalog rather than worksheet cell storage.
    pub fn tab<'e, 's>(&'e mut self, selector: impl Into<Selector<'s>>) -> Result<Option<Tab<'e>>> {
        guard::no_removal(self, "transaction")?;
        let tab = Snapshot::new(&self.base).tab(selector)?;
        Ok(tab.map(|tab| Tab {
            edit: self,
            position: tab.position(),
        }))
    }

    /// Move one tab immediately before another by semantic selector.
    ///
    /// `Ok(None)` means either selector did not resolve in the source
    /// snapshot. Names are the ordinary entry point; checked numeric selectors
    /// remain available for import and positional workflows.
    pub fn move_before<'a, 'b>(
        &mut self,
        sheet: impl Into<Selector<'a>>,
        anchor: impl Into<Selector<'b>>,
    ) -> Result<Option<&mut Self>> {
        guard::no_removal(self, "transaction")?;
        let Some((sheet, anchor)) = Snapshot::new(&self.base).pair(sheet, anchor)? else {
            return Ok(None);
        };
        order::move_relative(self, sheet.position(), anchor.position(), false)?;
        Ok(Some(self))
    }

    /// Move one tab immediately after another by semantic selector.
    ///
    /// `Ok(None)` means either selector did not resolve in the source snapshot.
    pub fn move_after<'a, 'b>(
        &mut self,
        sheet: impl Into<Selector<'a>>,
        anchor: impl Into<Selector<'b>>,
    ) -> Result<Option<&mut Self>> {
        guard::no_removal(self, "transaction")?;
        let Some((sheet, anchor)) = Snapshot::new(&self.base).pair(sheet, anchor)? else {
            return Ok(None);
        };
        order::move_relative(self, sheet.position(), anchor.position(), true)?;
        Ok(Some(self))
    }

    /// Move a selected tab to a checked zero-based final position.
    ///
    /// `Ok(None)` means the source selector or destination position does not
    /// exist. Prefer [`Self::move_before`] and [`Self::move_after`] when a
    /// stable semantic anchor is available.
    pub fn move_to<'a>(
        &mut self,
        sheet: impl Into<Selector<'a>>,
        position: usize,
    ) -> Result<Option<&mut Self>> {
        guard::no_removal(self, "transaction")?;
        let Some(sheet) = Snapshot::new(&self.base).tab(sheet)? else {
            return Ok(None);
        };
        if position >= self.base.len() {
            return Ok(None);
        }
        order::move_to(self, sheet.position(), position)?;
        Ok(Some(self))
    }

    /// Copy direct row and column page breaks between two worksheets.
    ///
    /// Both selectors are resolved against the exact source snapshot. The
    /// copied value is staged as owned semantic state, so publication never
    /// aliases worksheet XML or relationship identifiers.
    ///
    /// # Errors
    ///
    /// Returns an error when either selector targets a non-worksheet, source
    /// page breaks are invalid, or the transaction contains a removal plan.
    pub fn copy_page_breaks<'source, 'target>(
        &mut self,
        source: impl Into<Selector<'source>>,
        target: impl Into<Selector<'target>>,
    ) -> Result<Option<&mut Self>> {
        guard::no_removal(self, "page-break copy")?;
        let Some((source, target)) = Snapshot::new(&self.base).worksheet_pair(source, target)?
        else {
            return Ok(None);
        };
        let value = self.pending_page_breaks(&source)?;
        self.sheets
            .entry(target.position())
            .or_default()
            .page_breaks = Some(value);
        Ok(Some(self))
    }

    /// Move direct row and column page breaks between two worksheets.
    ///
    /// The source is cleared and the target receives an owned copy in one
    /// atomic workbook transaction. Moving a worksheet to itself is a no-op.
    ///
    /// # Errors
    ///
    /// Returns an error under the same dependency and validation conditions as
    /// [`Self::copy_page_breaks`].
    pub fn move_page_breaks<'source, 'target>(
        &mut self,
        source: impl Into<Selector<'source>>,
        target: impl Into<Selector<'target>>,
    ) -> Result<Option<&mut Self>> {
        guard::no_removal(self, "page-break move")?;
        let Some((source, target)) = Snapshot::new(&self.base).worksheet_pair(source, target)?
        else {
            return Ok(None);
        };
        if source.position() == target.position() {
            return Ok(Some(self));
        }
        let value = self.pending_page_breaks(&source)?;
        self.sheets
            .entry(target.position())
            .or_default()
            .page_breaks = Some(value);
        self.sheets
            .entry(source.position())
            .or_default()
            .page_breaks = Some(crate::page_breaks::PageBreaks::new());
        Ok(Some(self))
    }

    /// Create or replace one worksheet's complete page-margin set.
    pub fn put_page_margins<'a>(
        &mut self,
        sheet: impl Into<Selector<'a>>,
        value: crate::page_margins::Margins,
    ) -> Result<Option<&mut Self>> {
        guard::no_removal(self, "page margins")?;
        let Some(sheet) = Snapshot::new(&self.base).worksheet(sheet)? else {
            return Ok(None);
        };
        self.sheets
            .entry(sheet.position())
            .or_default()
            .page_margins = Some(OptionalAction::Put(value));
        Ok(Some(self))
    }

    /// Remove one worksheet's direct page margins.
    pub fn remove_page_margins<'a>(
        &mut self,
        sheet: impl Into<Selector<'a>>,
    ) -> Result<Option<&mut Self>> {
        guard::no_removal(self, "page margins")?;
        let Some(sheet) = Snapshot::new(&self.base).worksheet(sheet)? else {
            return Ok(None);
        };
        self.sheets
            .entry(sheet.position())
            .or_default()
            .page_margins = Some(OptionalAction::Remove);
        Ok(Some(self))
    }

    /// Copy exact typed page margins between two worksheets.
    pub fn copy_page_margins<'source, 'target>(
        &mut self,
        source: impl Into<Selector<'source>>,
        target: impl Into<Selector<'target>>,
    ) -> Result<Option<&mut Self>> {
        guard::no_removal(self, "page-margin copy")?;
        let Some((source, target)) = Snapshot::new(&self.base).worksheet_pair(source, target)?
        else {
            return Ok(None);
        };
        let value = self.pending_page_margins(&source)?;
        self.sheets
            .entry(target.position())
            .or_default()
            .page_margins = Some(OptionalAction::from_option(value));
        Ok(Some(self))
    }

    /// Move exact typed page margins between two worksheets atomically.
    pub fn move_page_margins<'source, 'target>(
        &mut self,
        source: impl Into<Selector<'source>>,
        target: impl Into<Selector<'target>>,
    ) -> Result<Option<&mut Self>> {
        guard::no_removal(self, "page-margin move")?;
        let Some((source, target)) = Snapshot::new(&self.base).worksheet_pair(source, target)?
        else {
            return Ok(None);
        };
        if source.position() == target.position() {
            return Ok(Some(self));
        }
        let value = self.pending_page_margins(&source)?;
        self.sheets
            .entry(target.position())
            .or_default()
            .page_margins = Some(OptionalAction::from_option(value));
        self.sheets
            .entry(source.position())
            .or_default()
            .page_margins = Some(OptionalAction::Remove);
        Ok(Some(self))
    }

    /// Create or replace one worksheet's complete relationship-free page setup.
    pub fn put_page_setup<'a>(
        &mut self,
        sheet: impl Into<Selector<'a>>,
        value: crate::page_setup::Setup,
    ) -> Result<Option<&mut Self>> {
        guard::no_removal(self, "page setup")?;
        let Some(sheet) = Snapshot::new(&self.base).worksheet(sheet)? else {
            return Ok(None);
        };
        self.sheets.entry(sheet.position()).or_default().page_setup =
            Some(OptionalAction::Put(value));
        Ok(Some(self))
    }

    /// Remove one worksheet's direct page setup when no printer settings are attached.
    pub fn remove_page_setup<'a>(
        &mut self,
        sheet: impl Into<Selector<'a>>,
    ) -> Result<Option<&mut Self>> {
        guard::no_removal(self, "page setup")?;
        let Some(sheet) = Snapshot::new(&self.base).worksheet(sheet)? else {
            return Ok(None);
        };
        self.sheets.entry(sheet.position()).or_default().page_setup = Some(OptionalAction::Remove);
        Ok(Some(self))
    }

    /// Copy the exact typed page setup between two worksheets.
    pub fn copy_page_setup<'source, 'target>(
        &mut self,
        source: impl Into<Selector<'source>>,
        target: impl Into<Selector<'target>>,
    ) -> Result<Option<&mut Self>> {
        guard::no_removal(self, "page-setup copy")?;
        let Some((source, target)) = Snapshot::new(&self.base).worksheet_pair(source, target)?
        else {
            return Ok(None);
        };
        let value = self.pending_page_setup(&source)?;
        self.sheets.entry(target.position()).or_default().page_setup =
            Some(OptionalAction::from_option(value));
        Ok(Some(self))
    }

    /// Move the exact typed page setup between two worksheets atomically.
    pub fn move_page_setup<'source, 'target>(
        &mut self,
        source: impl Into<Selector<'source>>,
        target: impl Into<Selector<'target>>,
    ) -> Result<Option<&mut Self>> {
        guard::no_removal(self, "page-setup move")?;
        let Some((source, target)) = Snapshot::new(&self.base).worksheet_pair(source, target)?
        else {
            return Ok(None);
        };
        if source.position() == target.position() {
            return Ok(Some(self));
        }
        let value = self.pending_page_setup(&source)?;
        self.sheets.entry(target.position()).or_default().page_setup =
            Some(OptionalAction::from_option(value));
        self.sheets.entry(source.position()).or_default().page_setup = Some(OptionalAction::Remove);
        Ok(Some(self))
    }

    /// Create or replace one worksheet's direct print options.
    pub fn put_print_options<'a>(
        &mut self,
        sheet: impl Into<Selector<'a>>,
        value: crate::print_options::PrintOptions,
    ) -> Result<Option<&mut Self>> {
        guard::no_removal(self, "print options")?;
        let Some(sheet) = Snapshot::new(&self.base).worksheet(sheet)? else {
            return Ok(None);
        };
        self.sheets
            .entry(sheet.position())
            .or_default()
            .print_options = Some(OptionalAction::Put(value));
        Ok(Some(self))
    }

    /// Remove one worksheet's direct print options.
    pub fn remove_print_options<'a>(
        &mut self,
        sheet: impl Into<Selector<'a>>,
    ) -> Result<Option<&mut Self>> {
        guard::no_removal(self, "print options")?;
        let Some(sheet) = Snapshot::new(&self.base).worksheet(sheet)? else {
            return Ok(None);
        };
        self.sheets
            .entry(sheet.position())
            .or_default()
            .print_options = Some(OptionalAction::Remove);
        Ok(Some(self))
    }

    /// Insert or replace one typed worksheet hyperlink by its checked range.
    pub fn put_hyperlink<'a>(
        &mut self,
        sheet: impl Into<Selector<'a>>,
        hyperlink: crate::hyperlinks::Hyperlink,
    ) -> Result<Option<&mut Self>> {
        guard::no_removal(self, "hyperlink")?;
        let Some(sheet) = Snapshot::new(&self.base).worksheet(sheet)? else {
            return Ok(None);
        };
        self.put_hyperlink_at(sheet.position(), hyperlink)?;
        Ok(Some(self))
    }

    /// Remove one typed worksheet hyperlink by its checked range.
    pub fn remove_hyperlink<'a>(
        &mut self,
        sheet: impl Into<Selector<'a>>,
        reference: crate::hyperlinks::HyperlinkReference,
    ) -> Result<Option<crate::hyperlinks::Hyperlink>> {
        guard::no_removal(self, "hyperlink")?;
        let Some(sheet) = Snapshot::new(&self.base).worksheet(sheet)? else {
            return Ok(None);
        };
        self.remove_hyperlink_at(sheet.position(), reference)
    }

    /// Copy exact typed print options between two worksheets.
    pub fn copy_print_options<'source, 'target>(
        &mut self,
        source: impl Into<Selector<'source>>,
        target: impl Into<Selector<'target>>,
    ) -> Result<Option<&mut Self>> {
        guard::no_removal(self, "print-options copy")?;
        let Some((source, target)) = Snapshot::new(&self.base).worksheet_pair(source, target)?
        else {
            return Ok(None);
        };
        let value = self.pending_print_options(&source)?;
        self.sheets
            .entry(target.position())
            .or_default()
            .print_options = Some(OptionalAction::from_option(value));
        Ok(Some(self))
    }

    /// Move exact typed print options between two worksheets atomically.
    pub fn move_print_options<'source, 'target>(
        &mut self,
        source: impl Into<Selector<'source>>,
        target: impl Into<Selector<'target>>,
    ) -> Result<Option<&mut Self>> {
        guard::no_removal(self, "print-options move")?;
        let Some((source, target)) = Snapshot::new(&self.base).worksheet_pair(source, target)?
        else {
            return Ok(None);
        };
        if source.position() == target.position() {
            return Ok(Some(self));
        }
        let value = self.pending_print_options(&source)?;
        self.sheets
            .entry(target.position())
            .or_default()
            .print_options = Some(OptionalAction::from_option(value));
        self.sheets
            .entry(source.position())
            .or_default()
            .print_options = Some(OptionalAction::Remove);
        Ok(Some(self))
    }

    /// Copy a bounded rectangular cell region with formula, text, and local
    /// shared-style dependencies closed inside this workbook.
    ///
    /// Scalar formulas are translated from each source coordinate to its
    /// corresponding target coordinate and their stale caches are discarded.
    /// Shared-string identity is retained exactly, including rich-text runs,
    /// while local style handles retain their validated workbook style-table
    /// lineage. Range-owned formulas, unknown cell encodings, and merged
    /// regions are refused rather than partially copied. An exact same-sheet,
    /// same-range transfer is a semantic no-op. Copying between distinct
    /// worksheets also clones selected one-/two-cell DrawingML anchors and
    /// their closed image or relationship-free classic-chart leaves into
    /// collision-free package parts. Existing target drawings, unsupported
    /// dependency graphs, and drawing-bearing moves are atomically refused
    /// rather than partially transferred.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid selectors, coordinates, ranges exceeding
    /// 65,536 cells, grid overflow, dependency scans exceeding their explicit
    /// bound, merged regions, unsupported formula/cell dependencies, or a
    /// worksheet that fails ordinary edit validation.
    pub fn copy_cells<'source, 'range, 'target, 'anchor>(
        &mut self,
        source: impl Into<Selector<'source>>,
        range: impl Into<Area<'range>>,
        target: impl Into<Selector<'target>>,
        anchor: impl Into<At<'anchor>>,
    ) -> Result<Option<&mut Self>> {
        self.transfer_cells(source, range, target, anchor, CellTransfer::Copy)
    }

    /// Copy a bounded dependency-free scalar range from an immutable donor
    /// workbook into this transaction's target workbook.
    ///
    /// Unlike [`Self::copy_cells`], this cross-workbook operation deliberately
    /// accepts only scalar cell values whose source worksheet has no modeled
    /// relationship or annotation surface. It never carries donor styles,
    /// shared-string identities, formulas, or package relationships across
    /// the workbook boundary. Target cells must be absent in the projected
    /// transaction state. Both workbooks must be ordinary, unsigned,
    /// plaintext, and unprotected; encrypted, signed, workbook-protected, or
    /// worksheet-protected donors are refused rather than declassified or
    /// copied through a lock. A donor from this exact workbook lineage
    /// delegates to [`Self::copy_cells`] so established same-workbook behavior
    /// remains unchanged.
    ///
    /// `Ok(None)` means either semantic sheet selector did not resolve.
    pub fn copy_cells_from<'source, 'range, 'target, 'anchor>(
        &mut self,
        donor: &Workbook,
        source: impl Into<Selector<'source>>,
        range: impl Into<Area<'range>>,
        target: impl Into<Selector<'target>>,
        anchor: impl Into<At<'anchor>>,
    ) -> Result<Option<&mut Self>> {
        guard::no_removal(self, "cross-workbook scalar cell transfer")?;
        if Arc::ptr_eq(&self.base.inner, &donor.inner) {
            return self.copy_cells(source, range, target, anchor);
        }
        super::super::transfer::copy_scalar_cells_from(self, donor, source, range, target, anchor)
    }

    /// Move a bounded cell region with the same dependency closure and
    /// refusals as [`Self::copy_cells`].
    ///
    /// Source-exclusive cells are removed in the same atomic transaction. For
    /// overlapping moves, destination writes are staged after source removals
    /// so the collected source snapshot wins deterministically.
    ///
    /// # Errors
    ///
    /// Returns the errors documented by [`Self::copy_cells`].
    pub fn move_cells<'source, 'range, 'target, 'anchor>(
        &mut self,
        source: impl Into<Selector<'source>>,
        range: impl Into<Area<'range>>,
        target: impl Into<Selector<'target>>,
        anchor: impl Into<At<'anchor>>,
    ) -> Result<Option<&mut Self>> {
        self.transfer_cells(source, range, target, anchor, CellTransfer::Move)
    }

    fn pending_transfer_styles(&self, position: usize, address: Address) -> TransferStyles {
        let Some(actions) = self.sheets.get(&position) else {
            return TransferStyles::default();
        };
        TransferStyles {
            font: actions.fonts.get(&address).cloned(),
            alignment: actions.alignments.get(&address).copied(),
            fill: actions.fills.get(&address).cloned(),
            border: actions.borders.get(&address).cloned(),
            protection: actions.protections.get(&address).copied(),
            quote_prefix: actions.quote_prefixes.get(&address).copied(),
            number_format: actions.number_formats.get(&address).cloned(),
        }
    }

    fn preflight_transfer_styles(
        &mut self,
        position: usize,
        address: Address,
        styles: &TransferStyles,
    ) -> Result<Vec<CaptureReservation>> {
        let mut reservations = Vec::with_capacity(7);
        let result = (|| -> Result<()> {
            let mut admit = |facet, bytes| -> Result<()> {
                reservations
                    .push(self.preflight_cell_capture_with_facet(position, address, facet, bytes)?);
                Ok(())
            };
            if let Some(change) = &styles.font {
                admit(CaptureFacet::Font, change.retained_bytes())?;
            }
            if let Some(change) = &styles.alignment {
                admit(CaptureFacet::Alignment, change.retained_bytes())?;
            }
            if let Some(change) = &styles.fill {
                admit(CaptureFacet::Fill, change.retained_bytes())?;
            }
            if let Some(change) = &styles.border {
                admit(CaptureFacet::Border, change.retained_bytes())?;
            }
            if let Some(change) = &styles.protection {
                admit(CaptureFacet::Protection, change.retained_bytes())?;
            }
            if styles.quote_prefix.is_some() {
                admit(CaptureFacet::QuotePrefix, size_of::<Option<bool>>())?;
            }
            if let Some(value) = &styles.number_format {
                admit(
                    CaptureFacet::NumberFormat,
                    size_of::<Option<Arc<str>>>()
                        .saturating_add(value.as_ref().map_or(0, |value| value.len())),
                )?;
            }
            Ok(())
        })();
        if let Err(error) = result {
            self.rollback_capture_reservations(&mut reservations);
            return Err(error);
        }
        Ok(reservations)
    }

    fn discard_transfer_styles(&mut self, position: usize, address: Address) -> Result<()> {
        self.discard_fill(position, address)?;
        self.discard_border(position, address)?;
        let actions = self.sheets.entry(position).or_default();
        actions.fonts.remove(&address);
        actions.alignments.remove(&address);
        actions.protections.remove(&address);
        actions.quote_prefixes.remove(&address);
        actions.number_formats.remove(&address);
        Ok(())
    }

    fn stage_transfer_styles(
        &mut self,
        position: usize,
        address: Address,
        styles: TransferStyles,
    ) -> Result<()> {
        if let Some(change) = styles.font {
            self.sheets
                .entry(position)
                .or_default()
                .fonts
                .insert(address, change);
        }
        if let Some(change) = styles.alignment {
            self.sheets
                .entry(position)
                .or_default()
                .alignments
                .insert(address, change);
        }
        if let Some(change) = styles.fill {
            self.stage_fill(position, address, &change)?;
        }
        if let Some(change) = styles.border {
            self.stage_border(position, address, &change)?;
        }
        if let Some(change) = styles.protection {
            self.sheets
                .entry(position)
                .or_default()
                .protections
                .insert(address, change);
        }
        if let Some(value) = styles.quote_prefix {
            self.sheets
                .entry(position)
                .or_default()
                .quote_prefixes
                .insert(address, value);
        }
        if let Some(value) = styles.number_format {
            self.sheets
                .entry(position)
                .or_default()
                .number_formats
                .insert(address, value);
        }
        Ok(())
    }

    /// Prove the aggregate deferred-style ledger before transfer mutation.
    /// `stage_fill` and `stage_border` are fallible because their counters are
    /// bounded resources; doing the same replacement arithmetic here keeps a
    /// late style admission from leaving cells, sheets, or drawings half
    /// mutated.
    fn preflight_transfer_style_memory(
        &self,
        mode: CellTransfer,
        source_position: usize,
        target_position: usize,
        staged: &[StagedTransfer],
    ) -> Result<()> {
        let maximum = style_staged_memory_limit(&self.base);
        let mut per_sheet = BTreeMap::<usize, usize>::new();
        let mut total = 0usize;
        for (position, actions) in &self.sheets {
            let bytes = actions
                .staged_fill_bytes
                .checked_add(actions.staged_border_bytes)
                .ok_or_else(|| invalid("worksheet style retained bytes overflow"))?;
            if bytes != 0 {
                per_sheet.insert(*position, bytes);
                total = total
                    .checked_add(bytes)
                    .ok_or_else(|| invalid("worksheet style retained bytes overflow"))?;
            }
        }

        let mut discarded = BTreeSet::<(usize, Address)>::new();
        for (source_address, target_address, _, _, source_noop, target_noop) in staged {
            if matches!(mode, CellTransfer::Move) && !source_noop {
                discarded.insert((source_position, *source_address));
            }
            if !target_noop {
                discarded.insert((target_position, *target_address));
            }
        }
        for (position, address) in discarded {
            let Some(actions) = self.sheets.get(&position) else {
                continue;
            };
            let fill = actions
                .fills
                .get(&address)
                .map(FillChange::retained_bytes)
                .unwrap_or(0);
            let border = usize::from(actions.borders.contains_key(&address))
                .saturating_mul(border_staged_entry_bytes());
            let removed = fill
                .checked_add(border)
                .ok_or_else(|| invalid("worksheet style retained bytes overflow"))?;
            if removed == 0 {
                continue;
            }
            let sheet_bytes = per_sheet.get(&position).copied().unwrap_or(0);
            let next_sheet = sheet_bytes
                .checked_sub(removed)
                .ok_or_else(|| invalid("worksheet style retained bytes underflow"))?;
            if next_sheet == 0 {
                per_sheet.remove(&position);
            } else {
                per_sheet.insert(position, next_sheet);
            }
            total = total
                .checked_sub(removed)
                .ok_or_else(|| invalid("worksheet style retained bytes underflow"))?;
        }

        for (_, target_address, _, styles, _, target_noop) in staged {
            if *target_noop {
                continue;
            }
            if let Some(fill) = &styles.fill {
                let bytes = fill.retained_bytes();
                enforce_style_staged_memory(bytes, maximum, "worksheet fill staged bytes")?;
                let sheet = per_sheet.entry(target_position).or_default();
                *sheet = sheet
                    .checked_add(bytes)
                    .ok_or_else(|| invalid("worksheet fill retained bytes overflow"))?;
                total = total
                    .checked_add(bytes)
                    .ok_or_else(|| invalid("worksheet style retained bytes overflow"))?;
                enforce_style_staged_memory(total, maximum, "worksheet style staged bytes")?;
            }
            if styles.border.is_some() {
                let bytes = border_staged_entry_bytes();
                let sheet = per_sheet.entry(target_position).or_default();
                *sheet = sheet
                    .checked_add(bytes)
                    .ok_or_else(|| invalid("worksheet border retained bytes overflow"))?;
                total = total
                    .checked_add(bytes)
                    .ok_or_else(|| invalid("worksheet style retained bytes overflow"))?;
                enforce_style_staged_memory(total, maximum, "worksheet style staged bytes")?;
            }
            // Keep the address in the loop so the replacement/discard set is
            // visibly tied to the exact target plan; it also guards future
            // changes from accidentally admitting a second target generation.
            let _ = target_address;
        }
        Ok(())
    }

    fn pending_transfer_rich_text(
        &self,
        sheet: &Worksheet,
        address: Address,
        cache: &mut HashMap<(usize, Address), Option<crate::RichText>>,
    ) -> Result<Option<crate::RichText>> {
        if let Some(action) = self
            .sheets
            .get(&sheet.position())
            .and_then(|actions| actions.cells.get(&address))
        {
            if matches!(action, Action::Remove) {
                return Ok(None);
            }
            match action.payload() {
                Some(Payload::Rich(rich)) => return Ok(Some((**rich).clone())),
                Some(_) => return Ok(None),
                None => {},
            }
        }
        let key = (sheet.position(), address);
        if let Some(value) = cache.get(&key) {
            return Ok(value.clone());
        }
        // The transfer planner seeds every requested base source and
        // destination address that has no staged payload through
        // `Worksheet::rich_text_many` before this lookup. Missing cache
        // entries therefore mean a non-rich cell, not a reason to restart an
        // XML scan for one address.
        Ok(None)
    }

    fn pending_transfer_overrides_base_payload(&self, position: usize, address: Address) -> bool {
        self.sheets
            .get(&position)
            .and_then(|actions| actions.cells.get(&address))
            .is_some_and(|action| matches!(action, Action::Remove) || action.payload().is_some())
    }

    fn transfer_cells<'source, 'range, 'target, 'anchor>(
        &mut self,
        source: impl Into<Selector<'source>>,
        range: impl Into<Area<'range>>,
        target: impl Into<Selector<'target>>,
        anchor: impl Into<At<'anchor>>,
        mode: CellTransfer,
    ) -> Result<Option<&mut Self>> {
        guard::no_removal(self, "cell transfer")?;
        let Some((source_sheet, target_sheet)) =
            Snapshot::new(&self.base).worksheet_pair(source, target)?
        else {
            return Ok(None);
        };
        let source_range = range.into().resolve()?;
        let target_start = anchor.into().resolve()?;
        let target_end_row = target_start
            .row()
            .get()
            .checked_add(source_range.rows())
            .ok_or_else(|| invalid("cell-transfer target row overflow"))?;
        let target_end_column = target_start
            .column()
            .get()
            .checked_add(source_range.columns())
            .ok_or_else(|| invalid("cell-transfer target column overflow"))?;
        let target_range = Rect::new(target_start, target_end_row, target_end_column)?;
        let cells = u64::from(source_range.rows())
            .checked_mul(u64::from(source_range.columns()))
            .ok_or_else(|| invalid("cell-transfer area overflow"))?;
        if cells > MAX_CELL_TRANSFER {
            return Err(Error::Limit {
                kind: super::super::LimitKind::TransferCells.into(),
                observed: cells,
                maximum: MAX_CELL_TRANSFER,
            });
        }
        if source_sheet.position() == target_sheet.position() && source_range == target_range {
            return Ok(Some(self));
        }
        self.ensure_unmerged_transfer(&source_sheet, source_range)?;
        self.ensure_unmerged_transfer(&target_sheet, target_range)?;
        if self.drawings.contains_key(&source_sheet.position())
            || self.drawings.contains_key(&target_sheet.position())
        {
            return Err(Error::Unsupported {
                feature: "composing cell transfers through a staged drawing graph",
            });
        }
        let drawing = super::super::drawing_transfer::plan(
            &self.base,
            &source_sheet,
            &target_sheet,
            source_range,
            target_start,
        )?;
        if drawing.is_some() && matches!(mode, CellTransfer::Move) {
            return Err(Error::Unsupported {
                feature: "moving cell ranges with drawing dependencies",
            });
        }
        self.ensure_no_range_formula_transfer(&source_sheet, source_range)?;
        self.ensure_no_range_formula_transfer(&target_sheet, target_range)?;

        let capacity = usize::try_from(cells)
            .map_err(|error| invalid(format!("cell-transfer count exceeds usize: {error}")))?;
        let mut staged = Vec::<StagedTransfer>::new();
        staged
            .try_reserve_exact(capacity)
            .map_err(|source| allocation("cell-transfer plan", source))?;

        // Prime the rich-text cache from one requested projection pass per
        // physical source part. Shared strings retain the existing first-use
        // rich-index census plus the requested payload pass; ordinary shared
        // transfers pay only that cached census and no payload materialization.
        // `Worksheet::rich_text` is deliberately exact for a single lookup,
        // but restarting its XML reader for every transfer address makes a
        // rich range quadratic. The bulk projection retains only requested
        // run values; staged payload actions are handled from the pending map
        // and never force a read of their superseded base bytes.
        let mut source_addresses = Vec::new();
        source_addresses
            .try_reserve_exact(capacity)
            .map_err(|source| allocation("cell-transfer source addresses", source))?;
        let mut target_addresses = Vec::new();
        target_addresses
            .try_reserve_exact(capacity)
            .map_err(|source| allocation("cell-transfer target addresses", source))?;
        for row_offset in 0..source_range.rows() {
            for column_offset in 0..source_range.columns() {
                let source_address = Address::at(
                    source_range.start().row().get() + row_offset,
                    source_range.start().column().get() + column_offset,
                )?;
                let target_address = Address::at(
                    target_start.row().get() + row_offset,
                    target_start.column().get() + column_offset,
                )?;
                if matches!(
                    self.pending_cell_state(&target_sheet, target_address)?,
                    State::Cell {
                        content: Cell::Unknown(_),
                        ..
                    }
                ) {
                    return Err(Error::Unsupported {
                        feature: "overwriting unmodeled cell encodings",
                    });
                }
                source_addresses.push(source_address);
                target_addresses.push(target_address);
            }
        }
        let mut source_rich_addresses = Vec::new();
        source_rich_addresses
            .try_reserve_exact(capacity)
            .map_err(|source| allocation("cell-transfer source rich addresses", source))?;
        let mut target_rich_addresses = Vec::new();
        target_rich_addresses
            .try_reserve_exact(capacity)
            .map_err(|source| allocation("cell-transfer target rich addresses", source))?;
        for (&source_address, &target_address) in
            source_addresses.iter().zip(target_addresses.iter())
        {
            if !self
                .pending_transfer_overrides_base_payload(source_sheet.position(), source_address)
            {
                source_rich_addresses.push(source_address);
            }
            if !self
                .pending_transfer_overrides_base_payload(target_sheet.position(), target_address)
            {
                target_rich_addresses.push(target_address);
            }
        }
        let mut rich_cache = HashMap::new();
        let combined = source_rich_addresses
            .len()
            .checked_add(target_rich_addresses.len())
            .ok_or_else(|| invalid("cell-transfer address count overflow"))?;
        rich_cache
            .try_reserve(combined)
            .map_err(|source| allocation("cell-transfer rich cache", source))?;
        if source_sheet.position() == target_sheet.position() && combined != 0 {
            let mut addresses = Vec::new();
            addresses
                .try_reserve_exact(combined)
                .map_err(|source| allocation("cell-transfer rich addresses", source))?;
            addresses.extend(source_rich_addresses.iter().copied());
            addresses.extend(target_rich_addresses.iter().copied());
            for (address, value) in source_sheet.rich_text_many(&addresses)? {
                rich_cache.insert((source_sheet.position(), address), value);
            }
        } else if source_sheet.position() != target_sheet.position() {
            if !source_rich_addresses.is_empty() {
                for (address, value) in source_sheet.rich_text_many(&source_rich_addresses)? {
                    rich_cache.insert((source_sheet.position(), address), value);
                }
            }
            if !target_rich_addresses.is_empty() {
                for (address, value) in target_sheet.rich_text_many(&target_rich_addresses)? {
                    rich_cache.insert((target_sheet.position(), address), value);
                }
            }
        }
        for (&source_address, &target_address) in
            source_addresses.iter().zip(target_addresses.iter())
        {
            let source_state = self.pending_cell_state(&source_sheet, source_address)?;
            let source_styles =
                self.pending_transfer_styles(source_sheet.position(), source_address);
            let target_state = self.pending_cell_state(&target_sheet, target_address)?;
            let target_styles =
                self.pending_transfer_styles(target_sheet.position(), target_address);
            let source_rich =
                self.pending_transfer_rich_text(&source_sheet, source_address, &mut rich_cache)?;
            let target_rich =
                self.pending_transfer_rich_text(&target_sheet, target_address, &mut rich_cache)?;
            let action = if let Some(rich) = source_rich.clone() {
                // A rich shared-string cell still owns the workbook's exact
                // shared-string dependency. Keep that identity in the
                // transfer action; only an inline source or a pending rich
                // payload should become a new inline rich value.
                if matches!(
                    &source_state,
                    State::Cell {
                        shared_string: Some(_),
                        ..
                    }
                ) {
                    transfer_action(source_state.clone(), source_address, target_address)?
                } else {
                    transfer_rich_action(rich, &source_state)
                }
            } else {
                transfer_action(source_state.clone(), source_address, target_address)?
            };
            let source_noop = matches!(&source_state, State::Missing)
                && self
                    .sheets
                    .get(&source_sheet.position())
                    .and_then(|actions| actions.cells.get(&source_address))
                    .is_none()
                && source_styles.is_empty();
            let rich_noop = match (&source_rich, &target_rich) {
                (Some(source), Some(target)) => source == target,
                (None, None) => true,
                _ => false,
            };
            let target_noop =
                source_state == target_state && source_styles == target_styles && rich_noop;
            staged.push((
                source_address,
                target_address,
                action,
                source_styles,
                source_noop,
                target_noop,
            ));
        }

        // A transfer replaces the destination cell's prior payload and all
        // authored local style facets. Release those old semantic charges
        // before admission so a replacement is charged for its final state,
        // not for both generations. Keep a budget snapshot so a rejected
        // plan restores the caller's ledger byte-for-byte.
        let capture_budget_before = self.capture_budget.clone();
        for (source_address, target_address, _, _, source_noop, target_noop) in &staged {
            if matches!(mode, CellTransfer::Move) && !source_noop {
                self.release_transfer_capture(source_sheet.position(), *source_address);
            }
            if !target_noop {
                self.release_transfer_capture(target_sheet.position(), *target_address);
            }
        }
        self.recompute_capture_budget();

        if let Err(error) = self.preflight_transfer_style_memory(
            mode,
            source_sheet.position(),
            target_sheet.position(),
            &staged,
        ) {
            self.capture_budget = capture_budget_before;
            return Err(error);
        }

        let mut capture_reservations = Vec::new();
        let admission = (|| -> Result<()> {
            for (source_address, target_address, action, styles, source_noop, target_noop) in
                &staged
            {
                if matches!(mode, CellTransfer::Move) && !source_noop {
                    capture_reservations.extend(self.preflight_transfer_action(
                        source_sheet.position(),
                        *source_address,
                        &Action::Remove,
                    )?);
                }
                if !target_noop {
                    capture_reservations.extend(self.preflight_transfer_action(
                        target_sheet.position(),
                        *target_address,
                        action,
                    )?);
                    capture_reservations.extend(self.preflight_transfer_styles(
                        target_sheet.position(),
                        *target_address,
                        styles,
                    )?);
                }
            }
            Ok(())
        })();
        if let Err(error) = admission {
            self.capture_budget = capture_budget_before.clone();
            return Err(error);
        }
        let mutation = (|| -> Result<()> {
            if matches!(mode, CellTransfer::Move) {
                for (source_address, _, _, _, source_noop, _) in &staged {
                    if !source_noop {
                        self.discard_transfer_styles(source_sheet.position(), *source_address)?;
                    }
                }
                let source_actions = self.actions(source_sheet.position());
                for (source_address, _, _, _, source_noop, _) in &staged {
                    if !source_noop {
                        source_actions.insert(*source_address, Action::Remove);
                    }
                }
            }
            for (_, target_address, action, styles, _, target_noop) in staged {
                if target_noop {
                    continue;
                }
                self.discard_transfer_styles(target_sheet.position(), target_address)?;
                let target_actions = self.actions(target_sheet.position());
                target_actions.insert(target_address, action);
                self.stage_transfer_styles(target_sheet.position(), target_address, styles)?;
            }
            if let Some(drawing) = drawing {
                // A picture-only transfer may have no effective cell action
                // at the target. Keep an empty sheet entry so commit visits
                // the target worksheet and publishes the drawing graph.
                self.sheets.entry(target_sheet.position()).or_default();
                self.drawings.insert(target_sheet.position(), drawing);
            }
            Ok(())
        })();
        if let Err(error) = mutation {
            self.rollback_capture_reservations(&mut capture_reservations);
            self.capture_budget = capture_budget_before;
            return Err(error);
        }
        Ok(Some(self))
    }

    /// Remove a worksheet selected by its developer-facing name or checked
    /// zero-based source position.
    ///
    /// `Ok(None)` means the selector did not resolve. The safe default refuses
    /// live formulas, unmodeled producer references, VBA projects, additional
    /// incoming relationships, and mixed mutation plans. Multiple independent
    /// worksheet removals may be collected in one atomic transaction.
    pub fn remove<'a>(&mut self, selector: impl Into<Selector<'a>>) -> Result<Option<&mut Self>> {
        let Some(sheet) = Snapshot::new(&self.base).worksheet(selector)? else {
            return Ok(None);
        };
        let position = sheet.position();
        let name = sheet.name().to_owned();
        if self.has_non_removal() {
            return Err(Error::SheetRemoveBlocked {
                sheet: name,
                position,
                part: "transaction".to_owned(),
                reason: RemoveBlock::MixedEdit,
            });
        }
        self.removed.insert(position);
        Ok(Some(self))
    }

    #[must_use]
    pub fn len(&self) -> usize {
        let existing = self.sheets.values().fold(
            self.removed
                .len()
                .saturating_add(usize::from(self.panes.is_some()))
                .saturating_add(usize::from(self.defined_names.is_some()))
                .saturating_add(self.drawings.len())
                .saturating_add(self.svg_lifecycle.len())
                .saturating_add(usize::from(self.active.is_some()))
                .saturating_add(
                    self.order
                        .as_ref()
                        .filter(|order| order.is_effective())
                        .map_or(0, |order| order.moves.len()),
                ),
            |len, actions| len.saturating_add(actions.len()),
        );
        self.added.iter().fold(existing, |len, added| {
            len.saturating_add(1)
                .saturating_add(usize::from(added.actions.defaults.is_some()))
                .saturating_add(usize::from(added.actions.web.is_some()))
                .saturating_add(added.actions.cells.len())
                .saturating_add(added.actions.fonts.len())
                .saturating_add(added.actions.alignments.len())
                .saturating_add(added.actions.fills.len())
                .saturating_add(added.actions.borders.len())
                .saturating_add(added.actions.protections.len())
                .saturating_add(added.actions.quote_prefixes.len())
                .saturating_add(added.actions.number_formats.len())
                .saturating_add(added.actions.rows.len())
                .saturating_add(added.actions.columns.len())
                .saturating_add(added.actions.merges.len())
                .saturating_add(usize::from(added.actions.page_breaks.is_some()))
                .saturating_add(usize::from(added.actions.page_margins.is_some()))
                .saturating_add(usize::from(added.actions.page_setup.is_some()))
                .saturating_add(usize::from(added.actions.print_options.is_some()))
                .saturating_add(added.actions.hyperlinks.len())
                .saturating_add(added.actions.form_control_scalars.len())
        })
    }

    pub fn is_empty(&self) -> bool {
        self.active.is_none()
            && self.panes.is_none()
            && self.defined_names.is_none()
            && self.drawings.is_empty()
            && self.svg_lifecycle.is_empty()
            && self
                .order
                .as_ref()
                .is_none_or(|order| !order.is_effective())
            && self.sheets.values().all(SheetActions::is_empty)
            && self.added.is_empty()
            && self.removed.is_empty()
    }

    /// Normalize deferred worksheet effects against the immutable source.
    ///
    /// Ordinary unbounded setters deliberately do not perform this census.
    /// A join and commit are the deferred boundaries: they build one effect
    /// plan, then apply it atomically so source-equivalent actions do not
    /// create conflicts, consume snapshot budget, or invalidate a signed
    /// no-op. Rich-text identity remains part of the comparison: a plain set
    /// over rich text is an authored representation change.
    fn normalize_effects(&mut self) -> Result<()> {
        self.normalize_effects_impl(false).map(|_| ())
    }

    fn normalize_effects_with_undo(&mut self) -> Result<NormalizationUndo> {
        self.normalize_effects_impl(true)
    }

    fn normalize_effects_impl(&mut self, retain_undo: bool) -> Result<NormalizationUndo> {
        let mut projections = Vec::new();
        for (&position, actions) in &self.sheets {
            let mut addresses = BTreeSet::new();
            addresses.extend(actions.cells.keys().copied());
            addresses.extend(actions.fonts.keys().copied());
            addresses.extend(actions.alignments.keys().copied());
            addresses.extend(actions.fills.keys().copied());
            addresses.extend(actions.borders.keys().copied());
            addresses.extend(actions.protections.keys().copied());
            addresses.extend(actions.quote_prefixes.keys().copied());
            addresses.extend(actions.number_formats.keys().copied());
            if addresses.is_empty() {
                continue;
            }
            let source = self
                .base
                .sheet(position)?
                .ok_or_else(|| invalid("worksheet effect target disappeared"))?;
            let store = source.store()?;
            for address in addresses {
                let stored = store.entry(address);
                let source_cell = stored.map(|value| &value.cell);
                let source_style = stored.and_then(|value| value.style);
                let source_format = source_style
                    .map(|style| self.base.inner.format(style))
                    .transpose()?;
                let original_action = actions.cells.get(&address);
                let format = match original_action {
                    Some(Action::Update {
                        style: Some(StyleEffect::Set(style)),
                        ..
                    }) => Some(self.base.inner.format(*style)?),
                    Some(Action::Update {
                        style: Some(StyleEffect::Reset),
                        ..
                    }) => Some(self.base.inner.format(0)?),
                    _ => source_format.clone(),
                };
                let font_matches = actions.fonts.get(&address).is_none_or(|change| {
                    font_change_matches_source(change, format.as_ref().and_then(Format::font))
                });
                let alignment_matches = actions.alignments.get(&address).is_none_or(|change| {
                    alignment_change_matches_source(
                        change,
                        format.as_ref().and_then(Format::alignment),
                    )
                });
                let fill_matches = actions.fills.get(&address).is_none_or(|change| {
                    fill_change_matches_source(change, format.as_ref().and_then(Format::fill))
                });
                let border_matches = actions.borders.get(&address).is_none_or(|change| {
                    border_change_matches_source(change, format.as_ref().and_then(Format::border))
                });
                let protection_matches = actions.protections.get(&address).is_none_or(|change| {
                    protection_change_matches_source(
                        change,
                        format.as_ref().and_then(Format::protection),
                    )
                });
                let quote_prefix_matches =
                    actions.quote_prefixes.get(&address).is_none_or(|value| {
                        quote_prefix_matches_source(
                            *value,
                            format.as_ref().map(Format::quote_prefix),
                        )
                    });
                let number_format_matches =
                    actions.number_formats.get(&address).is_none_or(|value| {
                        number_format_matches_source(
                            value.as_deref(),
                            format
                                .as_ref()
                                .and_then(|value| value.number_format().code()),
                        )
                    });
                let typed_effective = !(font_matches
                    && alignment_matches
                    && fill_matches
                    && border_matches
                    && protection_matches
                    && quote_prefix_matches
                    && number_format_matches);
                let payload_matches = match original_action.and_then(Action::payload) {
                    Some(payload) => payload_matches_source(&source, stored, address, payload)?,
                    None => true,
                };
                let style_matches = original_action
                    .and_then(|action| match action {
                        Action::Update { style, .. } => style.as_ref(),
                        Action::Remove => None,
                    })
                    .is_none_or(|style| style_matches_source(*style, source_style));
                let action = match original_action {
                    Some(Action::Update { .. })
                        if payload_matches && style_matches && !typed_effective =>
                    {
                        ActionNormalization::Remove
                    },
                    Some(Action::Remove) if typed_effective => ActionNormalization::RemoveToUpdate,
                    Some(Action::Remove) if source_cell.is_none() => ActionNormalization::Remove,
                    _ => ActionNormalization::Keep,
                };
                projections.push(NormalizationProjection {
                    position,
                    address,
                    payload_matches,
                    style_matches,
                    font: !font_matches,
                    alignment: !alignment_matches,
                    fill: !fill_matches,
                    border: !border_matches,
                    protection: !protection_matches,
                    quote_prefix: !quote_prefix_matches,
                    number_format: !number_format_matches,
                    action,
                });
            }
        }

        let mut plans = Vec::new();
        for projection in projections {
            let Some(actions) = self.sheets.get_mut(&projection.position) else {
                continue;
            };
            let mut plan = NormalizationPlan {
                position: projection.position,
                address: projection.address,
                action: None,
                font: None,
                alignment: None,
                fill: None,
                border: None,
                protection: None,
                quote_prefix: None,
                number_format: None,
            };
            if !projection.font {
                plan.font = actions.fonts.remove(&projection.address);
            }
            if !projection.alignment {
                plan.alignment = actions.alignments.remove(&projection.address);
            }
            if !projection.fill {
                plan.fill = actions.fills.remove(&projection.address);
            }
            if !projection.border {
                plan.border = actions.borders.remove(&projection.address);
            }
            if !projection.protection {
                plan.protection = actions.protections.remove(&projection.address);
            }
            if !projection.quote_prefix {
                plan.quote_prefix = actions.quote_prefixes.remove(&projection.address);
            }
            if !projection.number_format {
                plan.number_format = actions.number_formats.remove(&projection.address);
            }
            match projection.action {
                ActionNormalization::Remove => {
                    if let Some(action) = actions.cells.remove(&projection.address) {
                        plan.action = Some(ActionUndo::Removed(action));
                    }
                },
                ActionNormalization::RemoveToUpdate => {
                    if matches!(actions.cells.get(&projection.address), Some(Action::Remove)) {
                        actions.cells.insert(
                            projection.address,
                            Action::Update {
                                payload: None,
                                style: None,
                            },
                        );
                        plan.action = Some(ActionUndo::RemoveToUpdate);
                    }
                },
                ActionNormalization::Keep => {
                    let mut payload = None;
                    let mut style = None;
                    if let Some(Action::Update {
                        payload: current_payload,
                        style: current_style,
                    }) = actions.cells.get_mut(&projection.address)
                    {
                        if projection.payload_matches {
                            payload = current_payload.take();
                        }
                        if projection.style_matches {
                            style = current_style.take();
                        }
                    }
                    if payload.is_some() || style.is_some() {
                        plan.action = Some(ActionUndo::Update { payload, style });
                    }
                },
            }
            if plan.action.is_some()
                || plan.font.is_some()
                || plan.alignment.is_some()
                || plan.fill.is_some()
                || plan.border.is_some()
                || plan.protection.is_some()
                || plan.quote_prefix.is_some()
                || plan.number_format.is_some()
            {
                plans.push(plan);
            }
        }

        for position in self.sheets.keys().copied().collect::<Vec<_>>() {
            if let Some(actions) = self.sheets.get_mut(&position) {
                actions.staged_fill_bytes = actions
                    .fills
                    .values()
                    .map(FillChange::retained_bytes)
                    .fold(0, usize::saturating_add);
                actions.staged_border_bytes = actions
                    .borders
                    .keys()
                    .map(|_| border_staged_entry_bytes())
                    .fold(0, usize::saturating_add);
            }
        }

        let previous_capture_budget = self.capture_budget.take();
        let mut undo_capture_budget = None;
        if let Some(previous) = previous_capture_budget {
            match CaptureBudget::rebuilt(&previous, &self.base, &self.sheets, &self.added) {
                Ok(budget) => {
                    self.capture_budget = Some(budget);
                    if retain_undo {
                        undo_capture_budget = Some(previous);
                    }
                },
                Err(error) => {
                    NormalizationUndo {
                        plans,
                        capture_budget: Some(previous),
                    }
                    .restore(self);
                    return Err(error);
                },
            }
        }
        Ok(NormalizationUndo {
            plans,
            capture_budget: undo_capture_budget,
        })
    }

    /// Join an independently prepared edit when every effect is disjoint.
    ///
    /// Both edits must originate from the same immutable snapshot. On failure
    /// `self` is unchanged and [`JoinError`] returns ownership of `other`, so
    /// callers never lose prepared work while resolving a conflict.
    pub fn join(&mut self, other: Self) -> std::result::Result<&mut Self, JoinError> {
        if !Arc::ptr_eq(&self.base.inner, &other.base.inner) {
            return Err(JoinError {
                failure: JoinFailure::DifferentSnapshot,
                rejected: Box::new(other),
                normalization_source: None,
            });
        }
        self.join_in_place(other)?;
        Ok(self)
    }

    fn join_in_place(&mut self, other: Self) -> std::result::Result<(), JoinError> {
        let mut other = other;
        let mut normalization_undo = None;
        let mut self_normalization_undo = match self.normalize_effects_with_undo() {
            Ok(undo) => Some(undo),
            Err(error) => {
                return Err(JoinError {
                    failure: JoinFailure::Normalization(join_normalization(&error)),
                    rejected: Box::new(other),
                    normalization_source: Some(Box::new(error)),
                });
            },
        };
        macro_rules! reject {
            ($failure:expr) => {{
                if let Some(undo) = self_normalization_undo.take() {
                    undo.restore(self);
                }
                return Err(JoinError {
                    failure: $failure,
                    rejected: rejected_after_normalization(other, normalization_undo.take()),
                    normalization_source: None,
                });
            }};
        }
        macro_rules! reject_normalization {
            ($error:expr) => {{
                if let Some(undo) = self_normalization_undo.take() {
                    undo.restore(self);
                }
                let error = $error;
                return Err(JoinError {
                    failure: JoinFailure::Normalization(join_normalization(&error)),
                    rejected: rejected_after_normalization(other, normalization_undo.take()),
                    normalization_source: Some(Box::new(error)),
                });
            }};
        }
        if !Arc::ptr_eq(&self.base.inner, &other.base.inner) {
            reject!(JoinFailure::DifferentSnapshot);
        }
        // Normalize both sides at the deferred join boundary.  Setters on an
        // unbounded edit intentionally avoid source reads, but joining two
        // edits must compare their final effects rather than their authored
        // no-op requests. Both projections are compact and reversible, so a
        // source/format error cannot partially publish either edit.
        match other.normalize_effects_with_undo() {
            Ok(undo) => normalization_undo = Some(undo),
            Err(error) => {
                reject_normalization!(error);
            },
        }
        if self.panes.is_some() && other.panes.is_some() {
            reject!(JoinFailure::TaskPanes);
        }
        if (self.defined_names.is_some()
            && (other.defined_names.is_some() || other.has_defined_name_scope_changes()))
            || (other.defined_names.is_some() && self.has_defined_name_scope_changes())
        {
            reject!(JoinFailure::DefinedNames);
        }
        if self.drawings.keys().any(|position| {
            other.drawings.contains_key(position)
                || other
                    .drawings
                    .values()
                    .any(|plan| plan.source_position == *position)
        }) || other.drawings.keys().any(|position| {
            self.drawings
                .values()
                .any(|plan| plan.source_position == *position)
        }) {
            reject!(JoinFailure::DrawingTransfer);
        }
        if self.svg_lifecycle.iter().any(|intent| {
            other.svg_lifecycle.iter().any(|candidate| {
                intent.position == candidate.position && intent.selector == candidate.selector
            })
        }) {
            reject!(JoinFailure::Overlap(ConflictSet {
                conflicts: Box::new([]),
            }));
        }
        let conflicts = self.conflicts_with(&other);
        if !conflicts.is_empty() {
            reject!(JoinFailure::Overlap(conflicts));
        }
        let projected_fill_bytes = match projected_fill_retained_bytes(&self.sheets, &other.sheets)
        {
            Ok(bytes) => bytes,
            Err(_) => {
                reject!(JoinFailure::ResourceLimit(fill_join_resource_limit(
                    u64::MAX,
                    style_staged_memory_limit(&self.base),
                )));
            },
        };
        let projected_border_bytes =
            match projected_border_staged_bytes(&self.sheets, &other.sheets) {
                Ok(bytes) => bytes,
                Err(_) => {
                    reject!(JoinFailure::ResourceLimit(border_join_resource_limit(
                        u64::MAX,
                        style_staged_memory_limit(&self.base),
                    )));
                },
            };
        let maximum = style_staged_memory_limit(&self.base);
        if projected_fill_bytes > maximum {
            reject!(JoinFailure::ResourceLimit(fill_join_resource_limit(
                u64::try_from(projected_fill_bytes).unwrap_or(u64::MAX),
                maximum,
            )));
        }
        if projected_border_bytes > maximum {
            reject!(JoinFailure::ResourceLimit(border_join_resource_limit(
                u64::try_from(projected_border_bytes).unwrap_or(u64::MAX),
                maximum,
            )));
        }
        let projected_style_bytes = match projected_fill_bytes.checked_add(projected_border_bytes) {
            Some(bytes) => bytes,
            None => {
                reject!(JoinFailure::ResourceLimit(style_join_resource_limit(
                    u64::MAX,
                    maximum,
                )));
            },
        };
        if projected_style_bytes > maximum {
            reject!(JoinFailure::ResourceLimit(style_join_resource_limit(
                u64::try_from(projected_style_bytes).unwrap_or(u64::MAX),
                maximum,
            )));
        }

        let added_offset = self.added.len();
        enum CaptureJoinError {
            Resource(usize, usize),
            Normalization(Error),
        }
        let joined_capture_budget_result = match (&self.capture_budget, &other.capture_budget) {
            (Some(left), Some(right)) => left
                .joined(right, added_offset)
                .map(Some)
                .map_err(|(observed, maximum)| CaptureJoinError::Resource(observed, maximum)),
            (Some(left), None) => {
                let derived = CaptureBudget::from_staged(
                    &other.base,
                    left.maximum,
                    &other.sheets,
                    &other.added,
                )
                .map_err(CaptureJoinError::Normalization);
                match derived {
                    Ok(derived) => left.joined(&derived, added_offset).map(Some).map_err(
                        |(observed, maximum)| CaptureJoinError::Resource(observed, maximum),
                    ),
                    Err(error) => Err(error),
                }
            },
            (None, Some(right)) => {
                let derived = CaptureBudget::from_staged(
                    &self.base,
                    right.maximum,
                    &self.sheets,
                    &self.added,
                )
                .map_err(CaptureJoinError::Normalization);
                match derived {
                    Ok(derived) => derived.joined(right, added_offset).map(Some).map_err(
                        |(observed, maximum)| CaptureJoinError::Resource(observed, maximum),
                    ),
                    Err(error) => Err(error),
                }
            },
            (None, None) => Ok(None),
        };
        let joined_capture_budget = match joined_capture_budget_result {
            Ok(joined) => joined,
            Err(CaptureJoinError::Resource(observed, maximum)) => {
                reject!(JoinFailure::ResourceLimit(snapshot_join_resource_limit(
                    observed, maximum,
                )));
            },
            Err(CaptureJoinError::Normalization(error)) => reject_normalization!(error),
        };

        let other_cross_workbook_scalar = other.cross_workbook_scalar;
        let Some(combined_svg_intents) = self
            .svg_lifecycle
            .len()
            .checked_add(other.svg_lifecycle.len())
        else {
            reject!(JoinFailure::Overlap(ConflictSet {
                conflicts: Box::new([]),
            }));
        };
        if combined_svg_intents > super::super::svg_lifecycle::MAX_SVG_LIFECYCLE_INTENTS {
            reject!(JoinFailure::Overlap(ConflictSet {
                conflicts: Box::new([]),
            }));
        }
        let other_svg_staged_payload_bytes = other.svg_staged_payload_bytes;
        let Some(combined_svg_staged_payload_bytes) = self
            .svg_staged_payload_bytes
            .checked_add(other_svg_staged_payload_bytes)
        else {
            reject!(JoinFailure::Overlap(ConflictSet {
                conflicts: Box::new([]),
            }));
        };
        if svg_plan::check_payload_budget(&self.base, 0, combined_svg_staged_payload_bytes).is_err()
        {
            reject!(JoinFailure::Overlap(ConflictSet {
                conflicts: Box::new([]),
            }));
        }
        if self
            .svg_lifecycle
            .try_reserve(other.svg_lifecycle.len())
            .is_err()
            || self
                .svg_preflight
                .try_reserve(other.svg_preflight.len())
                .is_err()
            || self
                .svg_projected_owners
                .try_reserve(other.svg_projected_owners.len())
                .is_err()
            || self.added.try_reserve(other.added.len()).is_err()
        {
            reject!(JoinFailure::Overlap(ConflictSet {
                conflicts: Box::new([]),
            }));
        }
        if self.panes.is_none() {
            self.panes = other.panes;
        }
        if self.defined_names.is_none() {
            self.defined_names = other.defined_names;
        }
        self.drawings.extend(other.drawings);
        self.svg_lifecycle.extend(other.svg_lifecycle);
        self.svg_preflight.extend(other.svg_preflight);
        self.svg_projected_owners.extend(other.svg_projected_owners);
        self.svg_staged_payload_bytes = combined_svg_staged_payload_bytes;
        // Both transactions share one base, so either census describes it.
        self.svg_base_part_bytes = self.svg_base_part_bytes.or(other.svg_base_part_bytes);
        if self.active.is_none() {
            self.active = other.active.map(|target| match target {
                Target::Base(position) => Target::Base(position),
                Target::Added(index) => Target::Added(added_offset.saturating_add(index)),
            });
        }
        if self
            .order
            .as_ref()
            .is_none_or(|order| !order.is_effective())
        {
            self.order = other.order;
        }
        for (position, actions) in other.sheets {
            match self.sheets.entry(position) {
                Entry::Vacant(entry) => {
                    entry.insert(actions);
                },
                Entry::Occupied(entry) => {
                    let accepted = entry.into_mut();
                    if accepted.rename.is_none() {
                        accepted.rename = actions.rename;
                    }
                    if accepted.visibility.is_none() {
                        accepted.visibility = actions.visibility;
                    }
                    if accepted.web.is_none() {
                        accepted.web = actions.web;
                    }
                    match (accepted.defaults.as_mut(), actions.defaults) {
                        (None, defaults) => accepted.defaults = defaults,
                        (Some(accepted), Some(defaults)) => accepted.merge(defaults),
                        (Some(_), None) => {},
                    }
                    for (address, action) in actions.cells {
                        match accepted.cells.entry(address) {
                            Entry::Vacant(entry) => {
                                entry.insert(action);
                            },
                            Entry::Occupied(mut entry) => {
                                entry.get_mut().merge(action);
                            },
                        }
                    }
                    for (address, change) in actions.fonts {
                        match accepted.fonts.entry(address) {
                            Entry::Vacant(entry) => {
                                entry.insert(change);
                            },
                            Entry::Occupied(mut entry) => {
                                entry.get_mut().merge(&change);
                            },
                        }
                    }
                    for (address, change) in actions.alignments {
                        match accepted.alignments.entry(address) {
                            Entry::Vacant(entry) => {
                                entry.insert(change);
                            },
                            Entry::Occupied(mut entry) => {
                                entry.get_mut().merge(&change);
                            },
                        }
                    }
                    for (address, change) in actions.fills {
                        if let Some(current) = accepted.fills.get(&address).cloned() {
                            let previous = current.retained_bytes();
                            let mut merged = current;
                            merged.merge(&change);
                            let retained = merged.retained_bytes();
                            accepted.fills.insert(address, merged);
                            accepted.staged_fill_bytes = accepted
                                .staged_fill_bytes
                                .saturating_sub(previous)
                                .saturating_add(retained);
                        } else {
                            accepted.staged_fill_bytes = accepted
                                .staged_fill_bytes
                                .saturating_add(change.retained_bytes());
                            accepted.fills.insert(address, change);
                        }
                    }
                    for (address, change) in actions.borders {
                        let is_new = !accepted.borders.contains_key(&address);
                        match accepted.borders.entry(address) {
                            Entry::Vacant(entry) => {
                                entry.insert(change);
                            },
                            Entry::Occupied(mut entry) => {
                                entry.get_mut().merge(&change);
                            },
                        }
                        if is_new {
                            accepted.staged_border_bytes = accepted
                                .staged_border_bytes
                                .saturating_add(border_staged_entry_bytes());
                        }
                    }
                    for (address, change) in actions.protections {
                        match accepted.protections.entry(address) {
                            Entry::Vacant(entry) => {
                                entry.insert(change);
                            },
                            Entry::Occupied(mut entry) => {
                                entry.get_mut().merge(&change);
                            },
                        }
                    }
                    for (address, value) in actions.quote_prefixes {
                        match accepted.quote_prefixes.entry(address) {
                            Entry::Vacant(entry) => {
                                entry.insert(value);
                            },
                            Entry::Occupied(mut entry) => {
                                *entry.get_mut() = value;
                            },
                        }
                    }
                    for (address, value) in actions.number_formats {
                        match accepted.number_formats.entry(address) {
                            Entry::Vacant(entry) => {
                                entry.insert(value);
                            },
                            Entry::Occupied(mut entry) => {
                                *entry.get_mut() = value;
                            },
                        }
                    }
                    for (row, action) in actions.rows {
                        match accepted.rows.entry(row) {
                            Entry::Vacant(entry) => {
                                entry.insert(action);
                            },
                            Entry::Occupied(mut entry) => entry.get_mut().merge(action),
                        }
                    }
                    for (column, action) in actions.columns {
                        match accepted.columns.entry(column) {
                            Entry::Vacant(entry) => {
                                entry.insert(action);
                            },
                            Entry::Occupied(mut entry) => entry.get_mut().merge(action),
                        }
                    }
                    accepted.merges.extend(actions.merges);
                    if accepted.page_breaks.is_none() {
                        accepted.page_breaks = actions.page_breaks;
                    }
                    if accepted.page_margins.is_none() {
                        accepted.page_margins = actions.page_margins;
                    }
                    if accepted.page_setup.is_none() {
                        accepted.page_setup = actions.page_setup;
                    }
                    if accepted.print_options.is_none() {
                        accepted.print_options = actions.print_options;
                    }
                    accepted.hyperlinks.extend(actions.hyperlinks);
                    accepted
                        .form_control_scalars
                        .extend(actions.form_control_scalars);
                },
            }
        }
        self.added.extend(other.added);
        self.removed.extend(other.removed);
        self.cross_workbook_scalar |= other_cross_workbook_scalar;
        self.capture_budget = joined_capture_budget;
        Ok(())
    }

    /// Validate and atomically publish a new immutable snapshot.
    ///
    /// An edit that staged nothing returns its source snapshot, a signed one
    /// included (ADR 0006); a changed edit of a signed workbook is refused
    /// with [`Error::Signed`].
    pub fn commit(mut self) -> Result<Commit> {
        self.normalize_effects()?;
        if self.is_empty() {
            let budget = self.capture_budget.clone();
            let source = self.base.clone();
            let commit = Commit {
                package: crate::Package::from_workbook(self.base),
                patch: Patch {
                    source: Some(source.clone()),
                    target: Some(source),
                    ..Patch::default()
                },
            };
            if let Some(budget) = budget.as_ref() {
                budget.check_patch(&commit.patch)?;
            }
            return Ok(commit);
        }
        codec::ensure_unsigned(&self.base)?;
        if !self.removed.is_empty() {
            let budget = self.capture_budget.clone();
            let commit = package::commit_removals(self)?;
            if let Some(budget) = budget.as_ref() {
                budget.check_patch(&commit.patch)?;
            }
            return Ok(commit);
        }
        let Self {
            base,
            panes: requested_panes,
            defined_names: requested_defined_names,
            mut drawings,
            active: requested_active,
            order: requested_order,
            mut sheets,
            added,
            removed: _,
            cross_workbook_scalar,
            svg_lifecycle: svg_intents,
            svg_preflight: _,
            svg_projected_owners: _,
            svg_staged_payload_bytes: _,
            svg_base_part_bytes,
            mut capture_budget,
        } = self;
        let has_svg_lifecycle = !svg_intents.is_empty();
        ensure_defined_name_edit_is_composable(
            requested_defined_names.as_deref(),
            requested_order.as_ref(),
            &sheets,
            &added,
        )?;
        validate_web_edit(&base, requested_panes.as_ref(), &sheets, &added)?;
        let mut changes = Vec::new();
        let mut package_changes = Vec::new();
        let mut parts = Vec::new();
        let mut form_control_validations: Vec<(litchi_opc::PackURI, OrdinaryFormControlOverlay)> =
            Vec::new();
        let mut relationship_changes = Vec::new();
        let mut validated_worksheet_stores = Vec::new();
        let mut needs_recalculation = false;
        let mut drawing_graph = Vec::new();
        let mut svg_parts = Vec::new();
        let mut content_types = None;
        let lowering_changed_sheets = sheets.values().any(SheetActions::has_typed_style_requests);
        let (styles_part_changed, projected_style_count) =
            lower_style_requests(&base, &mut sheets, &mut parts)?;

        // Style lowering canonicalizes property-level requests into their
        // final cell actions. Rebuild the bounded ledger from that published
        // shape so an equal typed style, a canceled request, or a merged
        // replacement does not retain a stale pre-lowering charge. This is
        // only paid by bounded edits and is still metadata-only for deferred
        // worksheet parts.
        //
        // `commit` normalized the edit (and rebuilt the ledger from the same
        // `base`, `sheets` and `added`) just before lowering. Lowering only
        // changes `sheets` when some worksheet staged a typed style request;
        // without one, the rebuild here would reproduce the ledger it
        // already holds, so it is skipped (record 0377).
        if lowering_changed_sheets && let Some(previous) = capture_budget.take() {
            capture_budget = Some(CaptureBudget::rebuilt(&previous, &base, &sheets, &added)?);
        }

        let effective_renames = take_effective_renames(&base, &mut sheets)?;
        if let Some((position, _)) = effective_renames.first() {
            let data = base
                .inner
                .sheets
                .get(*position)
                .ok_or_else(|| invalid("renamed tab disappeared during edit"))?;
            package::ensure_reorder_supported(&base, &data.name, *position)?;
        }
        let rename_by_position = effective_renames
            .iter()
            .map(|(position, name)| (*position, name))
            .collect::<HashMap<_, _>>();
        let effective_order = requested_order.filter(OrderPlan::is_effective);
        if let Some(order) = &effective_order {
            package::validate_order_plan(order, base.inner.sheets.len())?;
        }
        let final_order =
            FinalOrder::plan(base.inner.sheets.len(), effective_order.as_ref(), &added)?;
        if let Some(first) = added.first() {
            let position = final_order
                .position(Target::Added(0))
                .ok_or_else(|| invalid("first created worksheet has no final position"))?;
            package::ensure_reorder_supported(&base, first.name.as_str(), position)?;
        }
        validate_final_sheet_names(&base, &final_order, &added, &rename_by_position)?;
        for (position, after) in &effective_renames {
            let before = base
                .inner
                .sheets
                .get(*position)
                .ok_or_else(|| invalid("renamed tab disappeared during patch creation"))?;
            changes.push(Change::Rename {
                position: *position,
                before: before.name.clone().into_boxed_str(),
                after: after.as_str().into(),
            });
        }

        if let Some(order) = &effective_order {
            let first = order
                .moves
                .first()
                .ok_or_else(|| invalid("effective tab order has no semantic move"))?;
            let data = base
                .inner
                .sheets
                .get(first.sheet)
                .ok_or_else(|| invalid("moved tab disappeared during edit"))?;
            package::ensure_reorder_supported(&base, &data.name, first.from)?;
            for moved in &order.moves {
                let data = base
                    .inner
                    .sheets
                    .get(moved.sheet)
                    .ok_or_else(|| invalid("moved tab disappeared during patch creation"))?;
                changes.push(Change::Move {
                    sheet: data.name.clone().into_boxed_str(),
                    from: moved.from,
                    to: moved.to,
                });
            }
        }

        let mut effective_tabs = Vec::new();
        for (position, requested) in &sheets {
            let Some(action) = requested.visibility else {
                continue;
            };
            let data =
                base.inner.sheets.get(*position).ok_or_else(|| {
                    invalid(format!("edited sheet position {position} disappeared"))
                })?;
            let after = action.visibility();
            if data.visibility == after {
                continue;
            }
            effective_tabs.push((*position, action));
        }

        let final_base_is_visible = |position: usize| {
            sheets
                .get(&position)
                .and_then(|requested| requested.visibility)
                .map_or_else(
                    || {
                        base.inner
                            .sheets
                            .get(position)
                            .is_some_and(|sheet| sheet.visibility == Visibility::Visible)
                    },
                    |action| action == TabAction::Show,
                )
        };
        let added_is_visible = |index: usize| {
            added.get(index).is_some_and(|sheet| {
                sheet
                    .actions
                    .visibility
                    .is_none_or(|action| action == TabAction::Show)
            })
        };
        let any_visible = (0..base.inner.sheets.len()).any(&final_base_is_visible)
            || (0..added.len()).any(added_is_visible);
        if !effective_tabs.is_empty() && !any_visible {
            let (position, _) = effective_tabs
                .iter()
                .find(|(_, action)| *action != TabAction::Show)
                .copied()
                .ok_or_else(|| invalid("tab visibility invariant failed without a hide action"))?;
            let data = base
                .inner
                .sheets
                .get(position)
                .ok_or_else(|| invalid("last visible tab disappeared during edit"))?;
            return Err(Error::TabEditBlocked {
                sheet: data.name.clone(),
                position,
                reason: TabEditBlock::LastVisibleTab,
            });
        }

        let final_position = |target: Target| final_order.position(target);
        let final_is_visible = |target: Target| match target {
            Target::Base(position) => final_base_is_visible(position),
            Target::Added(index) => added_is_visible(index),
        };

        let current_active = base.inner.active_sheet;
        let current_target = current_active.map(Target::Base);
        let final_active = if let Some(target) = requested_active {
            let name = match target {
                Target::Base(identity) => {
                    let data =
                        base.inner.sheets.get(identity).ok_or_else(|| {
                            invalid("requested active tab disappeared during edit")
                        })?;
                    data.name.as_str()
                },
                Target::Added(index) => {
                    let data = added.get(index).ok_or_else(|| {
                        invalid("requested new active tab disappeared during edit")
                    })?;
                    data.name.as_str()
                },
            };
            let position = final_order
                .position(target)
                .ok_or_else(|| invalid("requested active tab has no final position"))?;
            if !final_is_visible(target) {
                return Err(Error::TabEditBlocked {
                    sheet: name.to_owned(),
                    position,
                    reason: TabEditBlock::NotVisible,
                });
            }
            Some(target)
        } else if effective_tabs.is_empty() || current_target.is_some_and(final_is_visible) {
            current_target
        } else {
            let len = final_order.len();
            if len == 0 {
                None
            } else {
                current_target
                    .and_then(final_position)
                    .and_then(|current_position| {
                        (1..=len)
                            .filter_map(|offset| {
                                current_position.checked_add(offset).map(|sum| sum % len)
                            })
                            .filter_map(|position| final_order.target(position))
                            .find(|target| final_is_visible(*target))
                    })
                    .or_else(|| {
                        final_order
                            .targets
                            .iter()
                            .copied()
                            .find(|target| final_is_visible(*target))
                    })
            }
        };
        let final_active_position = final_active.and_then(final_position);
        if let Some(position) = final_active_position
            && position > raw::catalog_edit::MAX_ACTIVE_TAB
        {
            let target =
                final_active.ok_or_else(|| invalid("active position has no sheet identity"))?;
            let name = match target {
                Target::Base(identity) => base
                    .inner
                    .sheets
                    .get(identity)
                    .map(|sheet| sheet.name.as_str()),
                Target::Added(index) => added.get(index).map(|sheet| sheet.name.as_str()),
            }
            .ok_or_else(|| invalid("replacement active tab disappeared during edit"))?;
            return Err(Error::TabEditBlocked {
                sheet: name.to_owned(),
                position,
                reason: TabEditBlock::ActiveTabLimit,
            });
        }

        let active_before = current_active
            .map(|identity| package::active_tab_at(&base, identity, identity, None))
            .transpose()?;
        let active_after = final_active
            .zip(final_active_position)
            .map(|(target, position)| match target {
                Target::Base(identity) => package::active_tab_at(
                    &base,
                    identity,
                    position,
                    rename_by_position.get(&identity).map(|name| name.as_str()),
                ),
                Target::Added(index) => added
                    .get(index)
                    .map(|sheet| ActiveTab {
                        name: sheet.name.as_str().into(),
                        position,
                    })
                    .ok_or_else(|| invalid("new active tab disappeared during patch creation")),
            })
            .transpose()?;
        let active_change = (current_target.zip(current_active)
            != final_active.zip(final_active_position))
        .then_some(final_active_position)
        .flatten();
        if active_change.is_some() {
            let before = active_before
                .ok_or_else(|| invalid("non-empty workbook has no source active tab"))?;
            let after = active_after
                .ok_or_else(|| invalid("non-empty workbook has no final active tab"))?;
            changes.push(Change::Active { before, after });
        }

        for (position, action) in &effective_tabs {
            let data = base
                .inner
                .sheets
                .get(*position)
                .ok_or_else(|| invalid("effective tab disappeared during edit"))?;
            changes.push(Change::Visibility {
                sheet: data.name.clone().into_boxed_str(),
                position: *position,
                before: data.visibility.clone(),
                after: action.visibility(),
            });
        }

        for (position, requested) in sheets {
            let drawing = drawings.remove(&position);
            let data =
                base.inner.sheets.get(position).cloned().ok_or_else(|| {
                    invalid(format!("edited sheet position {position} disappeared"))
                })?;
            let SheetActions {
                rename: _,
                visibility: _,
                defaults,
                web,
                cells,
                fonts,
                alignments,
                fills,
                staged_fill_bytes: _,
                borders,
                staged_border_bytes: _,
                protections,
                quote_prefixes,
                number_formats,
                rows,
                columns,
                merges,
                page_breaks,
                page_margins,
                page_setup,
                print_options,
                hyperlinks,
                form_control_scalars,
            } = requested;
            if !form_control_scalars.is_empty() && data.kind != WorksheetKind::Worksheet {
                return Err(Error::NotWorksheet {
                    sheet: data.name.clone(),
                });
            }
            let form_control_overlays = if form_control_scalars.is_empty() {
                Vec::new()
            } else {
                crate::form_control::ordinary_form_control_overlays(
                    &base.inner.package,
                    &data.part_uri,
                    &form_control_scalars,
                )?
            };
            for overlay in &form_control_overlays {
                package_changes.push(PackageChange::FormControl {
                    sheet: data.name.as_str().into(),
                    control: overlay.position,
                    before: Box::new(crate::form_control::parse(
                        &overlay.before_properties,
                        &crate::form_control::Limits::default(),
                    )?),
                    after: Box::new(crate::form_control::parse(
                        &overlay.after_properties,
                        &crate::form_control::Limits::default(),
                    )?),
                });
            }
            form_control_validations.extend(
                form_control_overlays
                    .iter()
                    .cloned()
                    .map(|overlay| (data.part_uri.clone(), overlay)),
            );
            if defaults.is_none()
                && web.is_none()
                && cells.is_empty()
                && fonts.is_empty()
                && alignments.is_empty()
                && fills.is_empty()
                && borders.is_empty()
                && protections.is_empty()
                && quote_prefixes.is_empty()
                && number_formats.is_empty()
                && rows.is_empty()
                && columns.is_empty()
                && merges.is_empty()
                && page_breaks.is_none()
                && page_margins.is_none()
                && page_setup.is_none()
                && print_options.is_none()
                && hyperlinks.is_empty()
                && form_control_overlays.is_empty()
                && drawing.is_none()
            {
                continue;
            }
            if data.kind != WorksheetKind::Worksheet {
                return Err(Error::NotWorksheet {
                    sheet: data.name.clone(),
                });
            }
            let sheet = Worksheet {
                owner: Arc::clone(&base.inner),
                data: Arc::clone(&data),
                policy: None,
            };
            let store = sheet.store()?;
            let change_start = changes.len();
            let part = base.inner.package.get_part(&data.part_uri)?;
            let before = part.blob_arc();
            let effective_hyperlinks = project_hyperlink_actions(
                &data.name,
                &data.part_uri,
                &before,
                part.rels(),
                hyperlinks,
                &mut changes,
                &mut relationship_changes,
            )?;
            let effective_page_breaks = match page_breaks {
                Some(after) => {
                    let before = sheet.page_breaks()?;
                    if before == after {
                        None
                    } else {
                        changes.push(Change::PageBreaks {
                            sheet: data.name.clone().into_boxed_str(),
                            before,
                            after: after.clone(),
                        });
                        Some(after)
                    }
                },
                None => None,
            };
            let effective_page_margins = match page_margins {
                Some(action) => {
                    let before = sheet.page_margins()?;
                    let after = action.as_option().copied();
                    if before == after {
                        None
                    } else {
                        changes.push(Change::PageMargins {
                            sheet: data.name.clone().into_boxed_str(),
                            before,
                            after,
                        });
                        Some(action)
                    }
                },
                None => None,
            };
            let effective_page_setup = match page_setup {
                Some(action) => {
                    let before = sheet.page_setup()?;
                    let after = action.as_option().cloned();
                    if before == after {
                        None
                    } else {
                        changes.push(Change::PageSetup {
                            sheet: data.name.clone().into_boxed_str(),
                            before,
                            after: after.clone(),
                        });
                        Some(action)
                    }
                },
                None => None,
            };
            let effective_print_options = match print_options {
                Some(action) => {
                    let before = sheet.print_options()?;
                    let after = action.as_option().copied();
                    if before == after {
                        None
                    } else {
                        changes.push(Change::PrintOptions {
                            sheet: data.name.clone().into_boxed_str(),
                            before,
                            after,
                        });
                        Some(action)
                    }
                },
                None => None,
            };
            let effective_web = match web {
                Some(after) => {
                    let before = sheet.web_bindings()?.clone();
                    if before == after {
                        None
                    } else {
                        changes.push(Change::Web {
                            sheet: data.name.clone().into_boxed_str(),
                            before: before.clone(),
                            after: after.clone(),
                        });
                        Some(after)
                    }
                },
                None => None,
            };
            let merge_projection =
                project_merges(&data.name, Some(&sheet), Some(store), merges, &cells)?;
            for (range, change) in &merge_projection.changes {
                changes.push(Change::Merge {
                    sheet: data.name.clone().into_boxed_str(),
                    range: *range,
                    change: *change,
                });
            }
            let mut effective_defaults = None;
            if let Some(action) = defaults {
                let before = store.defaults().cloned();
                let after = defaults_after(before.as_ref(), action).map_err(|reason| {
                    Error::DefaultsEditBlocked {
                        sheet: data.name.clone(),
                        reason,
                    }
                })?;
                if before != after {
                    effective_defaults = Some(action);
                    changes.push(Change::Defaults {
                        sheet: data.name.clone().into_boxed_str(),
                        before,
                        after,
                    });
                }
            }
            let mut effective_cells = BTreeMap::new();
            for (address, action) in cells {
                let before = store.entry(address);
                if before.is_some_and(|stored| matches!(stored.cell, Cell::Unknown(_)))
                    && action.payload().is_some()
                {
                    return Err(Error::EditBlocked {
                        sheet: data.name.clone(),
                        address,
                        reason: EditBlock::UnknownCell,
                    });
                }
                let before_state = State::read(before, &base);
                let after_state = State::after(before, &action, &base);
                if before_state == after_state
                    && action_payload_matches_source(&sheet, before, address, &action)?
                {
                    continue;
                }
                needs_recalculation |= State::calculation_content(&before_state)
                    != State::calculation_content(&after_state);
                effective_cells.insert(address, action);
                changes.push(Change::Cell {
                    sheet: data.name.clone().into_boxed_str(),
                    address,
                    before: before_state,
                    after: after_state,
                });
            }
            let mut effective_rows = BTreeMap::new();
            for (row, action) in rows {
                let before = store.row_entry(row);
                let before_state = RowState::read(before, &base);
                let after_state = RowState::after(before, action, &base);
                if before_state == after_state {
                    continue;
                }
                effective_rows.insert(row, action);
                changes.push(Change::Row {
                    sheet: data.name.clone().into_boxed_str(),
                    row,
                    before: before_state,
                    after: after_state,
                });
            }
            let mut effective_columns = BTreeMap::new();
            for (column, action) in columns {
                let before = store.column_entry(column);
                let before_state = ColumnState::read(before, &base);
                let after_state = ColumnState::after(before, action, &base);
                if before_state == after_state {
                    continue;
                }
                effective_columns.insert(column, action);
                changes.push(Change::Column {
                    sheet: data.name.clone().into_boxed_str(),
                    column,
                    before: before_state,
                    after: after_state,
                });
            }
            if effective_defaults.is_none()
                && effective_web.is_none()
                && effective_cells.is_empty()
                && effective_rows.is_empty()
                && effective_columns.is_empty()
                && merge_projection.plan.is_empty()
                && effective_page_breaks.is_none()
                && effective_page_margins.is_none()
                && effective_page_setup.is_none()
                && effective_print_options.is_none()
                && effective_hyperlinks.is_none()
                && form_control_overlays.is_empty()
                && drawing.is_none()
            {
                continue;
            }

            let has_worksheet_edit = effective_defaults.is_some()
                || effective_web.is_some()
                || !effective_cells.is_empty()
                || !effective_rows.is_empty()
                || !effective_columns.is_empty()
                || !merge_projection.plan.is_empty()
                || effective_page_breaks.is_some()
                || effective_page_margins.is_some()
                || effective_page_setup.is_some()
                || effective_print_options.is_some()
                || effective_hyperlinks.is_some()
                || drawing.is_some();
            if !has_worksheet_edit {
                for overlay in form_control_overlays {
                    parts.push(PartChange {
                        uri: overlay.property_uri,
                        before: overlay.before_properties,
                        after: overlay.after_properties,
                        before_source: None,
                        after_source: None,
                    });
                    parts.push(PartChange {
                        uri: overlay.vml_uri,
                        before: overlay.before_vml,
                        after: overlay.after_vml,
                        before_source: None,
                        after_source: None,
                    });
                }
                continue;
            }

            let MergePlan { add, remove } = merge_projection.plan;
            let has_merge_removes = !remove.is_empty();
            let mut after = if !has_merge_removes {
                None
            } else {
                Some(raw::worksheet::edit::rewrite_merges(
                    &before,
                    &data.name,
                    MergePlan {
                        add: Vec::new(),
                        remove,
                    },
                )?)
            };
            let ordinary = Plan {
                defaults: effective_defaults,
                cells: effective_cells,
                rows: effective_rows,
                columns: effective_columns,
            };
            // A sparse payload-only edit on a large, plain worksheet can
            // validate its rewritten cells through the streaming worksheet
            // scanner. Keep the target list before ordinary takes ownership
            // of the cell map; the scanner retains only these records while
            // still validating the complete XML/MCE stream.
            let sparse_cell_targets = if ordinary.defaults.is_none()
                && ordinary.rows.is_empty()
                && ordinary.columns.is_empty()
                && effective_web.is_none()
                && effective_page_breaks.is_none()
                && effective_page_margins.is_none()
                && effective_page_setup.is_none()
                && effective_print_options.is_none()
                && effective_hyperlinks.is_none()
                && add.is_empty()
                && !has_merge_removes
                && drawing.is_none()
            {
                ordinary.cells.keys().copied().collect::<Vec<_>>()
            } else {
                Vec::new()
            };
            // Metadata-only worksheet edits already validate the source store
            // above and each metadata codec validates its own complete XML
            // surface. Keep the unchanged grid cold in the newly published
            // snapshot instead of parsing it again after the rewrite. Grid
            // and merge edits retain the full post-write store verification.
            let requires_store_verification =
                !ordinary.is_empty() || !add.is_empty() || has_merge_removes;
            // A value-only rewrite that is the worksheet's only byte change
            // may record the cell spans it copies verbatim from the source;
            // those records can narrow the post-write readback below. The
            // provenance writer emits exactly the ordinary writer's bytes.
            let value_only = after.is_none()
                && ordinary.defaults.is_none()
                && ordinary.rows.is_empty()
                && ordinary.columns.is_empty()
                && !ordinary.cells.is_empty()
                && add.is_empty()
                && effective_web.is_none()
                && effective_page_breaks.is_none()
                && effective_page_margins.is_none()
                && effective_page_setup.is_none()
                && effective_print_options.is_none()
                && effective_hyperlinks.is_none()
                && drawing.is_none();
            let changed_cells = ordinary.cells.len();
            let mut omitted: Box<[raw::worksheet::edit::OmittedCells]> = Box::new([]);
            if value_only {
                let rewrite = raw::worksheet::edit::rewrite_value_only_with_provenance(
                    &before,
                    &data.name,
                    ordinary.cells,
                    None,
                )?;
                omitted = rewrite.omitted;
                after = Some(rewrite.bytes);
            } else if !ordinary.is_empty() {
                let input = after.as_deref().unwrap_or(&before);
                after = Some(raw::worksheet::edit::rewrite(input, &data.name, ordinary)?);
            }
            if !add.is_empty() {
                let input = after.as_deref().unwrap_or(&before);
                after = Some(raw::worksheet::edit::rewrite_merges(
                    input,
                    &data.name,
                    MergePlan {
                        add,
                        remove: Vec::new(),
                    },
                )?);
            }
            if let Some(bindings) = &effective_web {
                let input = after.as_deref().unwrap_or(&before);
                after = Some(raw::web::replace(
                    input,
                    bindings,
                    &base.inner.package.read_limits(),
                )?);
            }
            if let Some(page_breaks) = &effective_page_breaks {
                let input = after.as_deref().unwrap_or(&before);
                after = Some(crate::page_breaks::replace(input, page_breaks)?);
            }
            if let Some(page_margins) = &effective_page_margins {
                let input = after.as_deref().unwrap_or(&before);
                after = Some(crate::page_margins::replace_page_margins(
                    input,
                    page_margins.as_option(),
                )?);
            }
            if let Some(page_setup) = &effective_page_setup {
                let input = after.as_deref().unwrap_or(&before);
                after = Some(crate::page_setup::replace_worksheet_page_setup(
                    input,
                    page_setup.as_option(),
                )?);
            }
            if let Some(print_options) = &effective_print_options {
                let input = after.as_deref().unwrap_or(&before);
                after = Some(crate::print_options::replace_print_options(
                    input,
                    print_options.as_option(),
                )?);
            }
            if let Some(hyperlinks) = &effective_hyperlinks {
                let input = after.as_deref().unwrap_or(&before);
                let relationship_ids = hyperlinks
                    .relationship_ids
                    .iter()
                    .map(Option::as_deref)
                    .collect::<Vec<_>>();
                let rewritten = crate::hyperlinks::codec::rewrite_hyperlinks(
                    input,
                    &hyperlinks.values,
                    &relationship_ids,
                )?;
                let parsed = crate::hyperlinks::codec::parse(
                    &rewritten,
                    &hyperlinks.verification_relationships,
                )?;
                if parsed != hyperlinks.values {
                    return Err(invalid(format!(
                        "worksheet hyperlink verification failed at {}",
                        data.name
                    )));
                }
                after = Some(rewritten);
            }
            if let Some(drawing) = drawing {
                let input = after.as_deref().unwrap_or(&before);
                after = Some(super::super::drawing_transfer::attach_worksheet(
                    input,
                    &drawing.target_relationship_id,
                )?);
                package_changes.push(PackageChange::DrawingTransfer {
                    source: drawing.source_name,
                    target: drawing.target_name,
                    anchors: drawing.anchors,
                    added: true,
                });
                drawing_graph.extend(drawing.graph);
            }
            let after =
                after.ok_or_else(|| invalid("effective worksheet edit produced no bytes"))?;
            let compacted =
                raw::compact::changed_worksheet(&after, "compact changed worksheet output")?;
            // Verification order: the reduced readback (change 0744) when it
            // is admitted; otherwise the sparse streaming verifier for a large
            // plain worksheet with a payload-only cell edit (spec-gap branch);
            // otherwise the complete parse, whose result and error remain
            // authoritative. The two narrowed readbacks each fall back rather
            // than refuse.
            let mut sparse_verified = false;
            let parsed = if requires_store_verification {
                let strings = || base.inner.shared_strings();
                match reduced_readback_store(
                    &after,
                    &compacted,
                    &omitted,
                    store.stored_cell_count(),
                    changed_cells,
                    strings,
                ) {
                    Some(reduced) => {
                        #[cfg(test)]
                        raw::worksheet::lane::route::note_admitted(
                            raw::worksheet::lane::route::Pass::Readback,
                        );
                        Some(VerifiedStore::Reduced(reduced))
                    },
                    None => {
                        sparse_verified = !sparse_cell_targets.is_empty()
                            && store.stored_cell_count() > MAX_VALIDATED_STORE_HANDOFF_CELLS
                            && store.merge_ranges().is_empty()
                            && verify_sparse_changed_cells(
                                compacted.bytes(),
                                &sparse_cell_targets,
                                &changes[change_start..],
                            )?;
                        if sparse_verified {
                            None
                        } else {
                            Some(VerifiedStore::Complete(raw::worksheet::parse(
                                compacted.bytes(),
                                strings,
                            )?))
                        }
                    },
                }
            } else {
                None
            };
            // Resolve web validation after grid parsing to preserve error order.
            // Compaction may already prove that an ordinary worksheet has no bindings.
            let (after, parsed_web) = compacted.into_bytes_and_web()?;
            if let Some(parsed) = parsed.as_ref() {
                if let Some(style_count) = projected_style_count {
                    validate_store_style_count(parsed.store(), style_count)?;
                } else {
                    base.inner.validate_styles(parsed.store())?;
                }
            }
            for change in &changes[change_start..] {
                match change {
                    Change::Create { .. }
                    | Change::Remove { .. }
                    | Change::Rename { .. }
                    | Change::Move { .. }
                    | Change::Active { .. }
                    | Change::Visibility { .. } => {},
                    Change::Web {
                        sheet,
                        after: expected,
                        ..
                    } => {
                        if &parsed_web != expected {
                            return Err(invalid(format!(
                                "worksheet web-binding verification failed at {sheet}"
                            )));
                        }
                    },
                    Change::Merge {
                        sheet,
                        range,
                        change,
                        ..
                    } => {
                        let parsed =
                            parsed.as_ref().map(VerifiedStore::store).ok_or_else(|| {
                                invalid("worksheet merge verification lost the parsed store")
                            })?;
                        if parsed.merge_ranges().contains(range) != change.after() {
                            return Err(invalid(format!(
                                "worksheet merged-range verification failed at {sheet}!{range}"
                            )));
                        }
                    },
                    Change::Defaults { sheet, after, .. } => {
                        let parsed =
                            parsed.as_ref().map(VerifiedStore::store).ok_or_else(|| {
                                invalid("worksheet defaults verification lost the parsed store")
                            })?;
                        if parsed.defaults() != after.as_ref() {
                            return Err(invalid(format!(
                                "worksheet defaults edit verification failed at {sheet}"
                            )));
                        }
                    },
                    Change::Cell {
                        sheet,
                        address,
                        after,
                        ..
                    } => {
                        if let Some(parsed) = parsed.as_ref().map(VerifiedStore::store) {
                            let actual = State::read(parsed.entry(*address), &base);
                            if actual != *after {
                                return Err(invalid(format!(
                                    "worksheet edit verification failed at {sheet}!{address}"
                                )));
                            }
                        } else if !sparse_verified {
                            return Err(invalid(
                                "worksheet cell verification lost the parsed store",
                            ));
                        }
                    },
                    Change::Row {
                        sheet, row, after, ..
                    } => {
                        let parsed =
                            parsed.as_ref().map(VerifiedStore::store).ok_or_else(|| {
                                invalid("worksheet row verification lost the parsed store")
                            })?;
                        let actual = RowState::read(parsed.row_entry(*row), &base);
                        if actual != *after {
                            return Err(invalid(format!(
                                "worksheet row edit verification failed at {sheet}!row {}",
                                row.get()
                            )));
                        }
                    },
                    Change::Column {
                        sheet,
                        column,
                        after,
                        ..
                    } => {
                        let parsed =
                            parsed.as_ref().map(VerifiedStore::store).ok_or_else(|| {
                                invalid("worksheet column verification lost the parsed store")
                            })?;
                        let actual = ColumnState::read(parsed.column_entry(*column), &base);
                        if actual != *after {
                            return Err(invalid(format!(
                                "worksheet column edit verification failed at {sheet}!column {}",
                                column.get()
                            )));
                        }
                    },
                    Change::PageBreaks {
                        sheet,
                        after: expected,
                        ..
                    } => {
                        let actual = crate::page_breaks::parse(&after)?;
                        if &actual != expected {
                            return Err(invalid(format!(
                                "worksheet page-break verification failed at {sheet}"
                            )));
                        }
                    },
                    Change::PageMargins {
                        sheet,
                        after: expected,
                        ..
                    } => {
                        let actual = crate::page_margins::parse_page_margins(&after)?;
                        if &actual != expected {
                            return Err(invalid(format!(
                                "worksheet page-margin verification failed at {sheet}"
                            )));
                        }
                    },
                    Change::PageSetup {
                        sheet,
                        after: expected,
                        ..
                    } => {
                        let actual = crate::page_setup::parse_worksheet_page_setup(&after)?;
                        if &actual != expected {
                            return Err(invalid(format!(
                                "worksheet page-setup verification failed at {sheet}"
                            )));
                        }
                    },
                    Change::PrintOptions {
                        sheet,
                        after: expected,
                        ..
                    } => {
                        let actual = crate::print_options::parse_print_options(&after)?;
                        if &actual != expected {
                            return Err(invalid(format!(
                                "worksheet print-options verification failed at {sheet}"
                            )));
                        }
                    },
                    Change::Hyperlink { .. } => {},
                }
            }
            let after = Arc::new(after);
            // Only a complete readback describes the published worksheet; a
            // reduced one is used exclusively above the handoff bounds.
            if let Some(VerifiedStore::Complete(parsed)) = parsed {
                if parsed.stored_cell_count() <= MAX_VALIDATED_STORE_HANDOFF_CELLS
                    && after.len() <= MAX_VALIDATED_STORE_HANDOFF_BYTES
                {
                    validated_worksheet_stores.push(ValidatedWorksheetStore {
                        uri: data.part_uri.clone(),
                        content: Arc::clone(&after),
                        store: parsed,
                    });
                }
            }
            parts.push(PartChange {
                uri: data.part_uri.clone(),
                before,
                after,
                before_source: None,
                after_source: None,
            });
            for overlay in form_control_overlays {
                parts.push(PartChange {
                    uri: overlay.property_uri,
                    before: overlay.before_properties,
                    after: overlay.after_properties,
                    before_source: None,
                    after_source: None,
                });
                parts.push(PartChange {
                    uri: overlay.vml_uri,
                    before: overlay.before_vml,
                    after: overlay.after_vml,
                    before_source: None,
                    after_source: None,
                });
            }
        }
        if !drawings.is_empty() {
            return Err(invalid("drawing transfer target disappeared during commit"));
        }

        let active_added = match final_active {
            Some(Target::Added(index)) => Some(index),
            Some(Target::Base(_)) | None => None,
        };
        let mut created = codec::create_sheets(
            &base,
            added,
            &final_order.added_positions,
            active_added,
            &mut changes,
            &mut needs_recalculation,
        )?;

        if final_active != current_target {
            if let Some(old_active) = current_active {
                let data =
                    base.inner.sheets.get(old_active).ok_or_else(|| {
                        invalid("previous active sheet disappeared during tab edit")
                    })?;
                if data.kind == WorksheetKind::Unknown {
                    return Err(Error::TabEditBlocked {
                        sheet: data.name.clone(),
                        position: old_active,
                        reason: TabEditBlock::MarkupCompatibility,
                    });
                }
                codec::compose_part(&mut parts, &base, &data.part_uri, |content| {
                    raw::sheet_view_edit::rewrite(
                        content,
                        false,
                        raw::sheet_view_edit::Context {
                            sheet: &data.name,
                            position: old_active,
                        },
                    )
                })?;
            }
            let new_active = final_active
                .ok_or_else(|| invalid("active tab disappeared during selection rewrite"))?;
            if let Target::Base(new_active) = new_active {
                let data = base
                    .inner
                    .sheets
                    .get(new_active)
                    .ok_or_else(|| invalid("new active sheet disappeared during tab edit"))?;
                if data.kind == WorksheetKind::Unknown {
                    return Err(Error::TabEditBlocked {
                        sheet: data.name.clone(),
                        position: new_active,
                        reason: TabEditBlock::MarkupCompatibility,
                    });
                }
                codec::compose_part(&mut parts, &base, &data.part_uri, |content| {
                    raw::sheet_view_edit::rewrite(
                        content,
                        true,
                        raw::sheet_view_edit::Context {
                            sheet: &data.name,
                            position: final_active_position.unwrap_or(new_active),
                        },
                    )
                })?;
            }
        }

        let reference_renames = effective_renames
            .iter()
            .map(|(position, after)| {
                let before = base.inner.sheets.get(*position).ok_or_else(|| {
                    invalid("renamed sheet disappeared during reference planning")
                })?;
                Ok(raw::reference_edit::Rename {
                    before: &before.name,
                    after: after.as_str(),
                })
            })
            .collect::<Result<Vec<_>>>()?;
        let sheet_titles = if effective_renames.is_empty() {
            Vec::new()
        } else {
            base.inner
                .sheets
                .iter()
                .map(|sheet| sheet.name.as_str())
                .collect::<Vec<_>>()
        };
        if let Some((position, _)) = effective_renames.first() {
            let sheet = base
                .inner
                .sheets
                .get(*position)
                .ok_or_else(|| invalid("rename context sheet disappeared"))?;
            let mut reference_parts = base
                .inner
                .package
                .iter_parts()
                .filter(|part| {
                    part.partname() != &base.inner.workbook_uri
                        && codec::reference_part_named(part.partname(), part.content_type())
                })
                .map(|part| {
                    (
                        part.partname().clone(),
                        part.content_type()
                            == litchi_opc::constants::content_type::OFC_EXTENDED_PROPERTIES,
                    )
                })
                .collect::<Vec<_>>();
            reference_parts.sort_unstable_by(|left, right| left.0.as_str().cmp(right.0.as_str()));
            for (uri, extended_properties) in reference_parts {
                let part_name = uri.to_string();
                let part_titles = if extended_properties {
                    sheet_titles.as_slice()
                } else {
                    &[]
                };
                codec::compose_part_optional(&mut parts, &base, &uri, |content| {
                    raw::reference_edit::rewrite(
                        content,
                        &reference_renames,
                        raw::reference_edit::Context {
                            sheet: &sheet.name,
                            position: *position,
                            part: &part_name,
                            sheet_titles: part_titles,
                        },
                    )
                })?;
            }
            for created_sheet in &mut created {
                let part_name = created_sheet.graph.part.partname().to_string();
                if let Some(after) = raw::reference_edit::rewrite(
                    created_sheet.graph.part.blob(),
                    &reference_renames,
                    raw::reference_edit::Context {
                        sheet: &sheet.name,
                        position: *position,
                        part: &part_name,
                        sheet_titles: &[],
                    },
                )? {
                    created_sheet.graph.part.set_blob(after);
                }
                let parsed = raw::worksheet::parse(created_sheet.graph.part.blob(), || {
                    base.inner.shared_strings()
                })?;
                if let Some(style_count) = projected_style_count {
                    validate_store_style_count(&parsed, style_count)?;
                } else {
                    base.inner.validate_styles(&parsed)?;
                }
                for change in &mut changes {
                    if let Change::Cell {
                        sheet,
                        address,
                        after,
                        ..
                    } = change
                        && sheet.as_ref() == created_sheet.name.as_str()
                    {
                        *after = State::read(parsed.entry(*address), &base);
                    }
                }
            }
            for change in &mut changes {
                let Change::Web { sheet, after, .. } = change else {
                    continue;
                };
                if let Some(data) = base
                    .inner
                    .sheets
                    .iter()
                    .find(|data| data.name.as_str() == sheet.as_ref())
                {
                    let content = parts
                        .iter()
                        .find(|part| part.uri == data.part_uri)
                        .map_or_else(
                            || base.inner.package.get_part(&data.part_uri).map(Part::blob),
                            |part| Ok(part.after.as_slice()),
                        )?;
                    *after = raw::web::read(content, &base.inner.package.read_limits())?;
                } else if let Some(created) = created
                    .iter()
                    .find(|created| created.name.as_str() == sheet.as_ref())
                {
                    *after = raw::web::read(
                        created.graph.part.blob(),
                        &base.inner.package.read_limits(),
                    )?;
                }
            }
        }
        if effective_order.is_some() || !created.is_empty() {
            let existing_titles = base
                .inner
                .sheets
                .iter()
                .enumerate()
                .map(|(position, sheet)| {
                    rename_by_position
                        .get(&position)
                        .map_or(sheet.name.as_str(), |name| name.as_str())
                })
                .collect::<Vec<_>>();
            let mut property_order = Vec::new();
            property_order
                .try_reserve_exact(final_order.len())
                .map_err(|source| allocation("extended-properties sheet order", source))?;
            for target in final_order.targets.iter().copied() {
                property_order.push(match target {
                    Target::Base(identity) => raw::properties_edit::Sheet::Existing(identity),
                    Target::Added(index) => {
                        let sheet = created.get(index).ok_or_else(|| {
                            invalid("created worksheet disappeared from property order")
                        })?;
                        raw::properties_edit::Sheet::New(sheet.name.as_str())
                    },
                });
            }
            let mut property_parts = base
                .inner
                .package
                .iter_parts()
                .filter(|part| {
                    part.content_type()
                        == litchi_opc::constants::content_type::OFC_EXTENDED_PROPERTIES
                })
                .map(|part| part.partname().clone())
                .collect::<Vec<_>>();
            property_parts.sort_unstable_by(|left, right| left.as_str().cmp(right.as_str()));
            for uri in property_parts {
                codec::compose_part_optional(&mut parts, &base, &uri, |content| {
                    raw::properties_edit::arrange_sheets(content, &existing_titles, &property_order)
                })?;
            }
        }

        let style_guard = (!styles_part_changed && changes.iter().any(Change::uses_shared_style))
            .then(|| {
                let uri = base
                    .inner
                    .styles_uri
                    .as_ref()
                    .ok_or_else(|| invalid("shared style state has no styles part"))?
                    .clone();
                let content = base.inner.package.get_part(&uri)?.blob_arc();
                Ok::<_, Error>(StyleGuard { uri, content })
            })
            .transpose()?;
        let calculation_graph = if needs_recalculation {
            codec::calculation_chain_removal(&base)?
        } else {
            Vec::new()
        };

        if !effective_renames.is_empty()
            || !effective_tabs.is_empty()
            || active_change.is_some()
            || effective_order.is_some()
            || !created.is_empty()
            || needs_recalculation
            || requested_defined_names.is_some()
        {
            let workbook_part = base.inner.package.get_part(&base.inner.workbook_uri)?;
            let before = workbook_part.blob_arc();
            let referenced_workbook = if let Some((position, _)) = effective_renames.first() {
                let sheet = base
                    .inner
                    .sheets
                    .get(*position)
                    .ok_or_else(|| invalid("rename workbook context disappeared"))?;
                raw::reference_edit::rewrite(
                    &before,
                    &reference_renames,
                    raw::reference_edit::Context {
                        sheet: &sheet.name,
                        position: *position,
                        part: base.inner.workbook_uri.as_str(),
                        sheet_titles: &[],
                    },
                )?
            } else {
                None
            };
            let workbook_input = referenced_workbook.as_deref().unwrap_or(&before);
            let scope_source = (effective_order.is_some() || !created.is_empty())
                .then(|| raw::parse_catalog(workbook_input, &base.inner.package.read_limits()))
                .transpose()?;
            let raw_order = effective_order
                .as_ref()
                .map(|order| {
                    let moved = order
                        .moves
                        .first()
                        .ok_or_else(|| invalid("effective tab order lost its move context"))?;
                    let data =
                        base.inner.sheets.get(moved.sheet).ok_or_else(|| {
                            invalid("tab reorder context disappeared during commit")
                        })?;
                    let relationship_ids = order
                        .positions
                        .iter()
                        .map(|identity| {
                            base.inner
                                .sheets
                                .get(*identity)
                                .map(|sheet| sheet.relationship_id.as_str())
                                .ok_or_else(|| {
                                    invalid("tab reorder target disappeared during commit")
                                })
                        })
                        .collect::<Result<Vec<_>>>()?;
                    Ok::<_, Error>(raw::catalog_edit::Order {
                        sheet: &data.name,
                        position: moved.from,
                        relationship_ids,
                        local_scopes: base
                            .inner
                            .defined_names
                            .iter()
                            .filter(|name| name.local_sheet_id.is_some())
                            .count(),
                    })
                })
                .transpose()?;
            let raw_renames = effective_renames
                .iter()
                .map(|(position, after)| {
                    let data =
                        base.inner.sheets.get(*position).ok_or_else(|| {
                            invalid("tab rename target disappeared during commit")
                        })?;
                    Ok(raw::catalog_edit::Rename {
                        sheet: &data.name,
                        position: *position,
                        relationship_id: &data.relationship_id,
                        name: after.as_str(),
                    })
                })
                .collect::<Result<Vec<_>>>()?;
            let has_existing_catalog_edit =
                !effective_renames.is_empty() || !effective_tabs.is_empty() || raw_order.is_some();
            let mut after = if has_existing_catalog_edit {
                let tabs = effective_tabs
                    .iter()
                    .map(|(position, action)| {
                        let data = base.inner.sheets.get(*position).ok_or_else(|| {
                            invalid("tab rewrite target disappeared during commit")
                        })?;
                        Ok(raw::catalog_edit::Tab {
                            sheet: &data.name,
                            position: *position,
                            relationship_id: &data.relationship_id,
                            state: action.raw(),
                        })
                    })
                    .collect::<Result<Vec<_>>>()?;
                raw::catalog_edit::rewrite(
                    workbook_input,
                    raw::catalog_edit::Plan {
                        tabs,
                        renames: raw_renames,
                        active: None,
                        order: raw_order,
                    },
                )?
            } else {
                workbook_input.to_vec()
            };
            for (index, sheet) in created.iter().enumerate() {
                let position = base
                    .inner
                    .sheets
                    .len()
                    .checked_add(index)
                    .ok_or_else(|| invalid("physical worksheet append position overflow"))?;
                after = raw::catalog_edit::append(
                    &after,
                    raw::catalog_edit::Create {
                        sheet: sheet.name.as_str(),
                        position,
                        sheet_id: sheet.sheet_id,
                        relationship_id: &sheet.relationship_id,
                        state: sheet.visibility.raw(),
                    },
                )?;
            }
            if !final_order.matches_appended(effective_order.as_ref()) {
                let mut relationship_ids = Vec::new();
                relationship_ids
                    .try_reserve_exact(final_order.len())
                    .map_err(|source| allocation("final catalog order", source))?;
                for target in final_order.targets.iter().copied() {
                    relationship_ids.push(match target {
                        Target::Base(identity) => base
                            .inner
                            .sheets
                            .get(identity)
                            .map(|sheet| sheet.relationship_id.as_str())
                            .ok_or_else(|| {
                                invalid("existing worksheet disappeared from final catalog order")
                            })?,
                        Target::Added(index) => created
                            .get(index)
                            .map(|sheet| sheet.relationship_id.as_str())
                            .ok_or_else(|| {
                                invalid("created worksheet disappeared from final catalog order")
                            })?,
                    });
                }
                let context = final_order
                    .targets
                    .first()
                    .copied()
                    .ok_or_else(|| invalid("final catalog order has no context sheet"))?;
                let context_name = match context {
                    Target::Base(identity) => base
                        .inner
                        .sheets
                        .get(identity)
                        .map(|sheet| sheet.name.as_str()),
                    Target::Added(index) => created.get(index).map(|sheet| sheet.name.as_str()),
                }
                .ok_or_else(|| invalid("final catalog context sheet disappeared"))?;
                after = raw::catalog_edit::rewrite(
                    &after,
                    raw::catalog_edit::Plan {
                        tabs: Vec::new(),
                        renames: Vec::new(),
                        active: None,
                        order: Some(raw::catalog_edit::Order {
                            sheet: context_name,
                            position: 0,
                            relationship_ids,
                            local_scopes: base
                                .inner
                                .defined_names
                                .iter()
                                .filter(|name| name.local_sheet_id.is_some())
                                .count(),
                        }),
                    },
                )?;
            }
            if let Some(position) = active_change {
                let target = final_active
                    .ok_or_else(|| invalid("active-tab rewrite has no final sheet identity"))?;
                let sheet = match target {
                    Target::Base(identity) => base
                        .inner
                        .sheets
                        .get(identity)
                        .map(|sheet| sheet.name.as_str()),
                    Target::Added(index) => created.get(index).map(|sheet| sheet.name.as_str()),
                }
                .ok_or_else(|| invalid("active-tab rewrite target disappeared"))?;
                after = raw::catalog_edit::rewrite(
                    &after,
                    raw::catalog_edit::Plan {
                        tabs: Vec::new(),
                        renames: Vec::new(),
                        active: Some(raw::catalog_edit::Active { sheet, position }),
                        order: None,
                    },
                )?;
            }
            if needs_recalculation {
                after = raw::recalc::invalidate(&after)?;
            }
            if let Some(names) = &requested_defined_names {
                after = raw::catalog_edit::replace_defined_names(&after, names)?;
            }
            if after.as_slice() != before.as_slice() {
                after = raw::compact::changed(&after, "compact changed workbook output")?;
            }
            if has_existing_catalog_edit
                || !created.is_empty()
                || active_change.is_some()
                || requested_defined_names.is_some()
            {
                let catalog = raw::parse_catalog(&after, &base.inner.package.read_limits())?;
                if let Some(expected) = &requested_defined_names
                    && catalog.defined_names.as_slice() != expected.as_ref()
                {
                    return Err(invalid("workbook defined-name edit verification failed"));
                }
                if Some(catalog.active_sheet_index) != final_active_position {
                    return Err(invalid("workbook active-tab edit verification failed"));
                }
                if catalog.sheets.len() != final_order.len() {
                    return Err(invalid(
                        "workbook creation verification has the wrong sheet count",
                    ));
                }
                for (position, target) in final_order.targets.iter().copied().enumerate() {
                    let expected = match target {
                        Target::Base(identity) => base
                            .inner
                            .sheets
                            .get(identity)
                            .map(|sheet| sheet.relationship_id.as_str()),
                        Target::Added(index) => created
                            .get(index)
                            .map(|sheet| sheet.relationship_id.as_str()),
                    }
                    .ok_or_else(|| invalid("final worksheet order lost a sheet identity"))?;
                    let actual = catalog
                        .sheets
                        .get(position)
                        .ok_or_else(|| invalid("final worksheet order lost a catalog slot"))?;
                    if actual.relationship_id != expected {
                        return Err(invalid(format!(
                            "workbook order verification failed at position {position}"
                        )));
                    }
                }
                for (identity, action) in &effective_tabs {
                    let position = final_position(Target::Base(*identity)).ok_or_else(|| {
                        invalid("visible tab disappeared from the final workbook order")
                    })?;
                    let actual = catalog
                        .sheets
                        .get(position)
                        .ok_or_else(|| invalid("workbook tab edit verification lost a sheet"))?;
                    if !package::raw_visibility_matches(&actual.visibility, *action) {
                        let sheet = base
                            .inner
                            .sheets
                            .get(*identity)
                            .map_or("<missing sheet>", |sheet| sheet.name.as_str());
                        return Err(invalid(format!(
                            "workbook tab visibility verification failed at {sheet}"
                        )));
                    }
                }
                for (identity, expected_name) in &effective_renames {
                    let position = final_position(Target::Base(*identity)).ok_or_else(|| {
                        invalid("renamed tab disappeared from the final workbook order")
                    })?;
                    let actual = catalog
                        .sheets
                        .get(position)
                        .ok_or_else(|| invalid("workbook rename verification lost a sheet"))?;
                    let expected = base.inner.sheets.get(*identity).ok_or_else(|| {
                        invalid("renamed tab identity disappeared during verification")
                    })?;
                    if actual.relationship_id != expected.relationship_id
                        || actual.name != expected_name.as_str()
                    {
                        return Err(invalid(format!(
                            "workbook tab rename verification failed at position {position}"
                        )));
                    }
                }
                if effective_order.is_some() || !created.is_empty() {
                    let source = scope_source.as_ref().ok_or_else(|| {
                        invalid("defined-name scope source disappeared during verification")
                    })?;
                    package::verify_defined_name_scopes(
                        source,
                        &catalog,
                        base.inner.sheets.len(),
                        &final_order.targets,
                    )?;
                }
                for sheet in &created {
                    let actual = catalog.sheets.get(sheet.position).ok_or_else(|| {
                        invalid("created worksheet disappeared during catalog verification")
                    })?;
                    if actual.name != sheet.name.as_str()
                        || actual.relationship_id != sheet.relationship_id
                        || actual.sheet_id != sheet.sheet_id
                        || !package::raw_visibility_matches(&actual.visibility, sheet.visibility)
                    {
                        return Err(invalid(format!(
                            "workbook creation verification failed at position {}",
                            sheet.position
                        )));
                    }
                }
            }
            if after.as_slice() != before.as_slice() {
                parts.push(PartChange {
                    uri: base.inner.workbook_uri.clone(),
                    before,
                    after: Arc::new(after),
                    before_source: None,
                    after_source: None,
                });
            }
        }

        if let Some(after) = requested_defined_names
            && after.as_ref() != base.inner.defined_names.as_ref()
        {
            package_changes.push(PackageChange::DefinedNames {
                before: base.inner.defined_names.clone(),
                after,
            });
        }

        let mut graph = Vec::new();
        graph
            .try_reserve(
                created
                    .len()
                    .saturating_add(calculation_graph.len())
                    .saturating_add(drawing_graph.len()),
            )
            .map_err(|source| allocation("package graph changes", source))?;
        graph.extend(created.into_iter().map(|sheet| sheet.graph));
        graph.extend(calculation_graph);
        graph.extend(drawing_graph);

        let preexisting_relationship_changes = relationship_changes.clone();
        svg_plan::plan(
            &base,
            &svg_intents,
            svg_base_part_bytes,
            &graph,
            &preexisting_relationship_changes,
            &mut parts,
            &mut relationship_changes,
            &mut svg_parts,
            &mut content_types,
            &mut package_changes,
        )?;

        let web_patch = match requested_panes {
            Some(PanesAction::Put { panes, conformance }) => {
                let after = panes.clone();
                let patch = common_web::plan_put(
                    &base.inner.package,
                    panes,
                    conformance,
                    &common_web::Limits::standard(),
                )?;
                if patch.is_noop() {
                    None
                } else {
                    package_changes.push(PackageChange::TaskPanes {
                        before: base.task_panes()?.cloned(),
                        after: Some(after),
                    });
                    Some(patch)
                }
            },
            Some(PanesAction::Remove) => {
                let patch =
                    common_web::plan_remove(&base.inner.package, &common_web::Limits::standard())?;
                if patch.is_noop() {
                    None
                } else {
                    package_changes.push(PackageChange::TaskPanes {
                        before: base.task_panes()?.cloned(),
                        after: None,
                    });
                    Some(patch)
                }
            },
            None => None,
        };

        if changes.is_empty()
            && package_changes.is_empty()
            && parts.is_empty()
            && relationship_changes.is_empty()
            && content_types.is_none()
            && svg_parts.is_empty()
            && graph.is_empty()
            && web_patch.is_none()
        {
            let source = base.clone();
            let commit = Commit {
                package: crate::Package::from_workbook(base),
                patch: Patch {
                    source: Some(source.clone()),
                    target: Some(source),
                    ..Patch::default()
                },
            };
            if let Some(budget) = capture_budget.as_ref() {
                budget.check_patch(&commit.patch)?;
            }
            return Ok(commit);
        }
        let mut package = base.inner.package.clone();
        for change in &relationship_changes {
            change.validate(&package)?;
        }
        for change in &svg_parts {
            if change.action == super::super::model::GraphAction::Add {
                change.validate(&package)?;
            }
        }
        // Source-proven XML transitions are applied through their proofs so
        // stale provenance cannot silently fall back to a canonical rewrite.
        for change in &parts {
            if let (Some(expected), Some(replacement)) =
                (&change.before_source, &change.after_source)
            {
                if expected.bytes() != change.before.as_slice()
                    || replacement.bytes() != change.after.as_slice()
                {
                    return Err(Error::PatchConflict {
                        part: change.uri.to_string(),
                    });
                }
                package.try_replace_owned_xml_part(expected.bytes(), replacement.clone())?;
            } else if package.get_part(&change.uri)?.content_type()
                == crate::form_control::CONTROL_PROPERTIES_CONTENT_TYPE
            {
                package.try_replace_owned_xml_part_bytes(
                    &change.uri,
                    change.before.as_slice(),
                    Arc::clone(&change.after),
                )?;
            } else {
                package
                    .get_part_mut(&change.uri)?
                    .set_blob_shared(Arc::clone(&change.after));
            }
        }
        if let Some(change) = &content_types {
            change.validate(&package)?;
            change.apply(&mut package)?;
        }
        for change in &svg_parts {
            if change.action == super::super::model::GraphAction::Add {
                change.apply(&mut package)?;
            }
        }
        // Manifest edits must be validated before graph publication and are
        // applied after changed XML/relationship members are staged.
        // (The concrete SVG planner supplies this token when needed.)
        // Content types are otherwise untouched by ordinary worksheet edits.
        for change in &relationship_changes {
            change.apply(&mut package)?;
        }
        for change in &graph {
            change.validate(&package)?;
            change.apply(&mut package)?;
        }
        for change in &svg_parts {
            if change.action == super::super::model::GraphAction::Remove {
                change.validate(&package)?;
            }
        }
        // SVG media removals happen only after all graph transitions are
        // visible to the final incoming-edge census.
        for change in &svg_parts {
            if change.action == super::super::model::GraphAction::Remove {
                change.apply(&mut package)?;
            }
        }
        if let Some(web) = &web_patch {
            let commit = web.apply(&package)?;
            package = commit.into_package();
        }
        if has_svg_lifecycle {
            validate_svg_final_removals(&package, &svg_parts)?;
        }
        for (worksheet_uri, overlay) in &form_control_validations {
            crate::form_control::validate_ordinary_form_control_candidate(
                &package,
                worksheet_uri,
                overlay,
            )?;
        }
        let workbook = Workbook::from_package_with_styles(package, Some(&base))?;
        if needs_recalculation {
            let metadata = workbook.calculation_metadata()?;
            let properties = metadata.properties().ok_or_else(|| {
                invalid("formula invalidation removed workbook calculation properties")
            })?;
            if properties.calculation_id() != 0
                || !properties.full_calculation_on_load()
                || !properties.force_full_calculation()
                || properties.calculation_completed()
                || !properties.calculate_on_save()
            {
                return Err(invalid(
                    "formula invalidation did not survive compact workbook publication",
                ));
            }
        }
        if web_patch.is_some()
            || changes
                .iter()
                .any(|change| matches!(change, Change::Web { .. }))
        {
            codec::validate_web_integrity(&workbook)?;
        }
        workbook.adopt_validated_worksheet_stores(&base, validated_worksheet_stores)?;
        let svg_read_guards = if has_svg_lifecycle {
            svg_plan::read_guards(&base, &workbook, &svg_intents)?
        } else {
            Box::new([])
        };
        let svg_final_guards = if has_svg_lifecycle {
            svg_plan::capture_final_guards(&base, &workbook, &svg_intents)?
        } else {
            Box::new([])
        };
        // A stylesheet edit creates a fresh style lineage. Rebind both
        // semantic endpoints before publishing the patch so target states do
        // not retain keys owned by the source snapshot.
        for change in &mut changes {
            change.rebind_style(&workbook);
        }
        // A paired control edit depends on its complete owner graph, including
        // incoming edges. Pin this bounded profile to the validated immutable
        // source; matching only the two changed payloads is insufficient.
        let authority = (cross_workbook_scalar || !form_control_validations.is_empty())
            .then(|| PatchAuthority::new(base.clone(), workbook.clone()));
        let commit = Commit {
            package: crate::Package::from_workbook(workbook.clone()),
            patch: Patch {
                changes: changes.into_boxed_slice(),
                package_changes: package_changes.into_boxed_slice(),
                parts: parts.into_boxed_slice(),
                relationships: relationship_changes.into_boxed_slice(),
                content_types,
                svg_parts: svg_parts.into_boxed_slice(),
                graph: graph.into_boxed_slice(),
                web: web_patch,
                style_guard,
                source: Some(base),
                target: Some(workbook),
                authority,
                svg_read_guards,
                svg_final_guards,
            },
        };
        if let Some(budget) = capture_budget.as_ref() {
            budget.check_patch(&commit.patch)?;
        }
        Ok(commit)
    }

    fn add_placed<T>(&mut self, name: T, placement: Placement) -> Result<NewSheet<'_>>
    where
        T: TryInto<Name>,
        Error: From<T::Error>,
    {
        guard::no_removal(self, "transaction")?;
        let name = name.try_into().map_err(Error::from)?;
        let limit_position = self
            .base
            .len()
            .checked_add(self.added.len())
            .ok_or_else(|| invalid("worksheet position overflow"))?;
        if limit_position >= raw::catalog_edit::MAX_SHEETS {
            return Err(Error::TabEditBlocked {
                sheet: name.as_str().to_owned(),
                position: limit_position,
                reason: TabEditBlock::SheetLimit,
            });
        }
        let index = self.added.len();
        if let Some(budget) = &mut self.capture_budget {
            budget.charge_new_sheet(index, name.as_str().len())?;
        }
        self.added
            .try_reserve(1)
            .map_err(|source| allocation("worksheet creation", source))?;
        self.added.push(Added {
            name,
            actions: SheetActions::default(),
            placement,
        });
        let effective_order = self.order.as_ref().filter(|order| order.is_effective());
        let position = match FinalOrder::plan(self.base.len(), effective_order, &self.added)
            .and_then(|order| {
                order
                    .position(Target::Added(index))
                    .ok_or_else(|| invalid("new worksheet has no projected position"))
            }) {
            Ok(position) => position,
            Err(error) => {
                if self.added.pop().is_none() {
                    return Err(invalid(
                        "new worksheet planning failed and its rollback target disappeared",
                    ));
                }
                return Err(error);
            },
        };
        if self.added.get(index).is_none() {
            let _ = self.added.pop();
            return Err(invalid("worksheet creation plan disappeared"));
        }
        let added = self
            .added
            .get_mut(index)
            .ok_or_else(|| invalid("worksheet creation plan disappeared after validation"))?;
        let capture_budget = &mut self.capture_budget;
        Ok(NewSheet {
            added,
            active: &mut self.active,
            style_lineage: &self.base.inner.style_lineage,
            capture_budget,
            index,
            position,
        })
    }

    pub(super) fn actions(&mut self, position: usize) -> &mut BTreeMap<Address, Action> {
        &mut self.sheets.entry(position).or_default().cells
    }

    pub(super) fn stage_form_control_scalar<'a>(
        &mut self,
        position: usize,
        selector: impl Into<crate::form_control::ControlSelector<'a>>,
        field: crate::form_control::ScalarField,
        value: Option<crate::form_control::ScalarValue>,
    ) -> Result<()> {
        guard::no_removal(self, "form-control scalar edit")?;
        let data = self
            .base
            .inner
            .sheets
            .get(position)
            .ok_or_else(|| invalid("form-control worksheet disappeared"))?;
        if data.kind != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: data.name.clone(),
            });
        }
        let operation_count = self
            .sheets
            .get(&position)
            .map_or(0, |actions| actions.form_control_scalars.len())
            .checked_add(1)
            .ok_or_else(|| invalid("form-control scalar operation count overflow"))?;
        crate::form_control::validate_scalar_operation_count(
            operation_count,
            crate::form_control::OwnerLimits::from_read_limits(
                self.base.inner.package.read_limits(),
            ),
        )?;
        let action = crate::form_control::FormControlScalarAction::new(selector, field, value);
        let mut combined = self
            .sheets
            .get(&position)
            .map_or_else(Vec::new, |actions| actions.form_control_scalars.clone());
        combined
            .try_reserve(1)
            .map_err(|source| allocation("form-control scalar edit plan", source))?;
        combined.push(action);
        let overlays = crate::form_control::ordinary_form_control_overlays(
            &self.base.inner.package,
            &data.part_uri,
            &combined,
        )?;
        let actions = self.sheets.entry(position).or_default();
        if overlays.is_empty() {
            actions.form_control_scalars.clear();
        } else {
            actions.form_control_scalars = combined;
        }
        Ok(())
    }

    pub(super) fn web_bindings(&mut self, position: usize) -> Result<&mut WebBindings> {
        let needs_source = self
            .sheets
            .get(&position)
            .is_none_or(|actions| actions.web.is_none());
        if needs_source {
            let data = self
                .base
                .inner
                .sheets
                .get(position)
                .cloned()
                .ok_or_else(|| invalid("web-binding target disappeared"))?;
            if data.kind != WorksheetKind::Worksheet {
                return Err(Error::NotWorksheet {
                    sheet: data.name.clone(),
                });
            }
            let sheet = Worksheet {
                owner: Arc::clone(&self.base.inner),
                data,
                policy: None,
            };
            let source = sheet.web_bindings()?.clone();
            self.sheets.entry(position).or_default().web = Some(source);
        }
        self.sheets
            .get_mut(&position)
            .and_then(|actions| actions.web.as_mut())
            .ok_or_else(|| invalid("web-binding edit initialization failed"))
    }

    pub(super) fn put_hyperlink_at(
        &mut self,
        position: usize,
        hyperlink: crate::hyperlinks::Hyperlink,
    ) -> Result<()> {
        let address = hyperlink.reference().range().start();
        let key = crate::hyperlinks::model::reference_key(hyperlink.reference().range());
        let already_staged = self
            .sheets
            .get(&position)
            .is_some_and(|actions| actions.hyperlinks.contains_key(&key));
        let staged_count = self
            .sheets
            .get(&position)
            .map_or(0, |actions| actions.hyperlinks.len());
        if !already_staged && staged_count >= MAX_HYPERLINK_EDITS {
            return Err(invalid(format!(
                "worksheet hyperlink edit exceeds the {MAX_HYPERLINK_EDITS}-operation limit"
            )));
        }
        let authored_bytes = hyperlink.reference().as_str().len()
            + hyperlink.location().map_or(0, str::len)
            + hyperlink.external_target().map_or(0, str::len)
            + hyperlink.display().map_or(0, str::len)
            + hyperlink.tooltip().map_or(0, str::len)
            + size_of::<crate::hyperlinks::Hyperlink>();
        let previous_bytes = self
            .sheets
            .get(&position)
            .and_then(|actions| actions.hyperlinks.get(&key))
            .map_or(0, hyperlink_action_retained_bytes);
        let source = self.source_hyperlink(position, key)?;
        if source.as_ref() == Some(&hyperlink) {
            self.sheets
                .entry(position)
                .or_default()
                .hyperlinks
                .remove(&key);
            self.recompute_relationship_capture(position, address);
            return Ok(());
        }
        let _capture =
            self.preflight_hyperlink_capture(position, address, previous_bytes, authored_bytes)?;
        self.sheets
            .entry(position)
            .or_default()
            .hyperlinks
            .insert(key, HyperlinkAction::Put(hyperlink));
        Ok(())
    }

    pub(super) fn remove_hyperlink_at(
        &mut self,
        position: usize,
        reference: crate::hyperlinks::HyperlinkReference,
    ) -> Result<Option<crate::hyperlinks::Hyperlink>> {
        let address = reference.range().start();
        let key = crate::hyperlinks::model::reference_key(reference.range());
        let current = self.pending_hyperlink(position, key)?;
        let Some(current) = current else {
            return Ok(None);
        };
        let source = self.source_hyperlink(position, key)?;
        let staged_put = self
            .sheets
            .get(&position)
            .and_then(|actions| actions.hyperlinks.get(&key))
            .is_some_and(|action| matches!(action, HyperlinkAction::Put(_)));
        if staged_put {
            self.sheets
                .entry(position)
                .or_default()
                .hyperlinks
                .remove(&key);
            self.recompute_relationship_capture(position, address);
            return Ok(Some(current));
        }
        let previous_bytes = self
            .sheets
            .get(&position)
            .and_then(|actions| actions.hyperlinks.get(&key))
            .map_or(0, hyperlink_action_retained_bytes);
        if source.is_none() {
            let _capture =
                self.preflight_hyperlink_capture(position, address, previous_bytes, 0)?;
            self.sheets
                .entry(position)
                .or_default()
                .hyperlinks
                .remove(&key);
            self.recompute_relationship_capture(position, address);
        } else {
            let already_staged = self
                .sheets
                .get(&position)
                .is_some_and(|actions| actions.hyperlinks.contains_key(&key));
            let staged_count = self
                .sheets
                .get(&position)
                .map_or(0, |actions| actions.hyperlinks.len());
            if !already_staged && staged_count >= MAX_HYPERLINK_EDITS {
                return Err(invalid(format!(
                    "worksheet hyperlink edit exceeds the {MAX_HYPERLINK_EDITS}-operation limit"
                )));
            }
            let _capture = self.preflight_hyperlink_capture(
                position,
                address,
                previous_bytes,
                reference.as_str().len() + size_of::<crate::hyperlinks::Hyperlink>(),
            )?;
            self.sheets
                .entry(position)
                .or_default()
                .hyperlinks
                .insert(key, HyperlinkAction::Remove(reference));
        }
        Ok(Some(current))
    }

    fn pending_hyperlink(
        &self,
        position: usize,
        key: (u32, u32, u32, u32),
    ) -> Result<Option<crate::hyperlinks::Hyperlink>> {
        if let Some(action) = self
            .sheets
            .get(&position)
            .and_then(|actions| actions.hyperlinks.get(&key))
        {
            return Ok(match action {
                HyperlinkAction::Put(value) => Some(value.clone()),
                HyperlinkAction::Remove(_) => None,
            });
        }
        self.source_hyperlink(position, key)
    }

    fn source_hyperlink(
        &self,
        position: usize,
        key: (u32, u32, u32, u32),
    ) -> Result<Option<crate::hyperlinks::Hyperlink>> {
        let data = self
            .base
            .inner
            .sheets
            .get(position)
            .cloned()
            .ok_or_else(|| invalid("hyperlink target disappeared"))?;
        if data.kind != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: data.name.clone(),
            });
        }
        let sheet = Worksheet {
            owner: Arc::clone(&self.base.inner),
            data,
            policy: None,
        };
        Ok(sheet.hyperlinks()?.into_iter().find(|value| {
            crate::hyperlinks::model::reference_key(value.reference().range()) == key
        }))
    }

    fn pending_page_breaks(&self, sheet: &Worksheet) -> Result<crate::page_breaks::PageBreaks> {
        if let Some(value) = self
            .sheets
            .get(&sheet.position())
            .and_then(|actions| actions.page_breaks.as_ref())
        {
            return Ok(value.clone());
        }
        sheet.page_breaks()
    }

    fn pending_page_setup(&self, sheet: &Worksheet) -> Result<Option<crate::page_setup::Setup>> {
        if let Some(value) = self
            .sheets
            .get(&sheet.position())
            .and_then(|actions| actions.page_setup.as_ref())
        {
            return Ok(value.as_option().cloned());
        }
        sheet.page_setup()
    }

    fn pending_page_margins(
        &self,
        sheet: &Worksheet,
    ) -> Result<Option<crate::page_margins::Margins>> {
        if let Some(value) = self
            .sheets
            .get(&sheet.position())
            .and_then(|actions| actions.page_margins.as_ref())
        {
            return Ok(value.as_option().copied());
        }
        sheet.page_margins()
    }

    fn pending_print_options(
        &self,
        sheet: &Worksheet,
    ) -> Result<Option<crate::print_options::PrintOptions>> {
        if let Some(value) = self
            .sheets
            .get(&sheet.position())
            .and_then(|actions| actions.print_options.as_ref())
        {
            return Ok(value.as_option().copied());
        }
        sheet.print_options()
    }

    fn pending_cell_state(&self, sheet: &Worksheet, address: Address) -> Result<State> {
        let stored = sheet.store()?.entry(address);
        Ok(self
            .sheets
            .get(&sheet.position())
            .and_then(|actions| actions.cells.get(&address))
            .map_or_else(
                || State::read(stored, &self.base),
                |action| State::after(stored, action, &self.base),
            ))
    }

    fn ensure_unmerged_transfer(&self, sheet: &Worksheet, range: Rect) -> Result<()> {
        let store = sheet.store()?;
        let intents = self
            .sheets
            .get(&sheet.position())
            .map_or(&[][..], |actions| actions.merges.as_slice());
        for row in range.start().row().get()..range.end().0 {
            for column in range.start().column().get()..range.end().1 {
                let address = Address::at(row, column)?;
                if pending_merge(store.merge_ranges(), intents, address).is_some() {
                    return Err(Error::EditBlocked {
                        sheet: sheet.name().to_owned(),
                        address,
                        reason: EditBlock::CoveredMerge,
                    });
                }
            }
        }
        Ok(())
    }

    fn ensure_no_range_formula_transfer(&self, sheet: &Worksheet, range: Rect) -> Result<()> {
        for (position, entry) in sheet.store()?.entries().iter().enumerate() {
            if position >= MAX_CELL_DEPENDENCY_SCAN {
                return Err(Error::Limit {
                    kind: super::super::LimitKind::TransferDependencyScan.into(),
                    observed: u64::try_from(position)
                        .unwrap_or(u64::MAX)
                        .saturating_add(1),
                    maximum: MAX_CELL_DEPENDENCY_SCAN as u64,
                });
            }
            let State::Cell {
                content: Cell::Formula(formula),
                ..
            } = self.pending_cell_state(sheet, entry.address)?
            else {
                continue;
            };
            let owned = match formula.kind() {
                FormulaKind::Scalar => continue,
                FormulaKind::Array { range } | FormulaKind::DataTable { range } => range
                    .as_ref()
                    .map(|range| Rect::from_a1(range.as_str()))
                    .transpose()?
                    .unwrap_or_else(|| Rect::single(entry.address)),
                FormulaKind::Unknown(_) => Rect::single(entry.address),
            };
            if rectangles_overlap(range, owned) {
                return Err(Error::Unsupported {
                    feature: "copying or overwriting range-owned formulas",
                });
            }
        }
        Ok(())
    }

    pub(super) fn set_visibility(&mut self, position: usize, action: TabAction) {
        self.sheets.entry(position).or_default().visibility = Some(action);
    }

    pub(super) fn set_name(&mut self, position: usize, name: Name) {
        self.sheets.entry(position).or_default().rename = Some(name);
    }

    pub(super) fn set_active(&mut self, position: usize) {
        self.active = Some(Target::Base(position));
    }

    fn conflicts_with(&self, other: &Self) -> ConflictSet {
        let mut conflicts = Vec::new();
        let removal_conflict = self
            .removed
            .intersection(&other.removed)
            .next()
            .copied()
            .or_else(|| {
                (!self.removed.is_empty() && other.has_non_removal())
                    .then(|| self.removed.iter().next().copied())
                    .flatten()
            })
            .or_else(|| {
                (!other.removed.is_empty() && self.has_non_removal())
                    .then(|| other.removed.iter().next().copied())
                    .flatten()
            });
        if let Some(position) = removal_conflict {
            let sheet = self
                .base
                .inner
                .sheets
                .get(position)
                .map_or("<missing sheet>", |sheet| sheet.name.as_str());
            conflicts.push(Conflict::Remove {
                sheet: sheet.into(),
                position,
            });
        }
        if let (Some(left), Some(_)) = (
            self.order.as_ref().filter(|order| order.is_effective()),
            other.order.as_ref().filter(|order| order.is_effective()),
        ) {
            let moved = left.moves.first();
            let position = moved.map_or(0, |moved| moved.from);
            let sheet = moved
                .and_then(|moved| self.base.inner.sheets.get(moved.sheet))
                .map_or("<missing sheet>", |sheet| sheet.name.as_str());
            conflicts.push(Conflict::Order {
                sheet: sheet.into(),
                position,
            });
        }
        if let (Some(target), Some(_)) = (self.active, other.active) {
            let (sheet, position) = self.target_context(target);
            conflicts.push(Conflict::Active {
                sheet: sheet.into(),
                position,
            });
        }
        for (left_index, left) in self.added.iter().enumerate() {
            if other
                .added
                .iter()
                .any(|right| right.name.identity_key() == left.name.identity_key())
                || other.sheets.values().any(|actions| {
                    actions
                        .rename
                        .as_ref()
                        .is_some_and(|name| name.identity_key() == left.name.identity_key())
                })
            {
                conflicts.push(Conflict::Name {
                    sheet: left.name.as_str().into(),
                    position: order::projected_position(self, Target::Added(left_index))
                        .unwrap_or_else(|| self.base.len().saturating_add(left_index)),
                });
            }
        }
        for (right_index, right) in other.added.iter().enumerate() {
            if self.sheets.values().any(|actions| {
                actions
                    .rename
                    .as_ref()
                    .is_some_and(|name| name.identity_key() == right.name.identity_key())
            }) {
                conflicts.push(Conflict::Name {
                    sheet: right.name.as_str().into(),
                    position: order::projected_position(other, Target::Added(right_index))
                        .unwrap_or_else(|| other.base.len().saturating_add(right_index)),
                });
            }
        }
        for (position, left) in &self.sheets {
            let Some(right) = other.sheets.get(position) else {
                continue;
            };
            let sheet = self
                .base
                .inner
                .sheets
                .get(*position)
                .map_or("<missing sheet>", |sheet| sheet.name.as_str());
            if left.rename.is_some() && right.rename.is_some() {
                conflicts.push(Conflict::Name {
                    sheet: sheet.into(),
                    position: *position,
                });
            }
            if left.visibility.is_some() && right.visibility.is_some() {
                conflicts.push(Conflict::Tab {
                    sheet: sheet.into(),
                    position: *position,
                });
            }
            if !left.form_control_scalars.is_empty() && !right.form_control_scalars.is_empty() {
                conflicts.push(Conflict::FormControls {
                    sheet: sheet.into(),
                    position: *position,
                });
            }
            if left.web.is_some() && right.web.is_some() {
                conflicts.push(Conflict::Web {
                    sheet: sheet.into(),
                    position: *position,
                });
            }
            if left.page_breaks.is_some() && right.page_breaks.is_some() {
                conflicts.push(Conflict::PageBreaks {
                    sheet: sheet.into(),
                    position: *position,
                });
            }
            if left.page_margins.is_some() && right.page_margins.is_some() {
                conflicts.push(Conflict::PageMargins {
                    sheet: sheet.into(),
                    position: *position,
                });
            }
            if left.page_setup.is_some() && right.page_setup.is_some() {
                conflicts.push(Conflict::PageSetup {
                    sheet: sheet.into(),
                    position: *position,
                });
            }
            if left.print_options.is_some() && right.print_options.is_some() {
                conflicts.push(Conflict::PrintOptions {
                    sheet: sheet.into(),
                    position: *position,
                });
            }
            let hyperlink_references = left
                .hyperlinks
                .iter()
                .filter(|&(key, _)| right.hyperlinks.contains_key(key))
                .map(|(_, action)| match action {
                    HyperlinkAction::Put(value) => value.reference().range(),
                    HyperlinkAction::Remove(reference) => reference.range(),
                })
                .collect::<Vec<_>>();
            if !hyperlink_references.is_empty() {
                conflicts.push(Conflict::Hyperlinks {
                    sheet: sheet.into(),
                    position: *position,
                    references: hyperlink_references.into_boxed_slice(),
                });
            }
            if let (Some(left), Some(right)) = (left.defaults, right.defaults) {
                let fields = left.fields() & right.fields();
                if left.overlaps(right) {
                    conflicts.push(Conflict::Defaults {
                        sheet: sheet.into(),
                        position: *position,
                        fields,
                    });
                }
            }
            let ranges = merge_conflicts(left, right);
            if !ranges.is_empty() {
                conflicts.push(Conflict::Merges {
                    sheet: sheet.into(),
                    position: *position,
                    ranges: ranges.into_boxed_slice(),
                });
            }
            let addresses =
                conflicts::overlapping_keys(&left.cells, &right.cells, |left, right| {
                    left.overlaps(right)
                });
            let mut addresses = addresses;
            for (address, left_font) in &left.fonts {
                if let Some(right_font) = right.fonts.get(address)
                    && left_font.overlaps(right_font)
                {
                    addresses.push(*address);
                }
                if right
                    .cells
                    .get(address)
                    .is_some_and(action_has_style_effect)
                {
                    addresses.push(*address);
                }
            }
            for address in right.fonts.keys() {
                if left.cells.get(address).is_some_and(action_has_style_effect) {
                    addresses.push(*address);
                }
            }
            for (address, left_alignment) in &left.alignments {
                if let Some(right_alignment) = right.alignments.get(address)
                    && left_alignment.overlaps(right_alignment)
                {
                    addresses.push(*address);
                }
                if right
                    .cells
                    .get(address)
                    .is_some_and(action_has_style_effect)
                {
                    addresses.push(*address);
                }
            }
            for address in right.alignments.keys() {
                if left.cells.get(address).is_some_and(action_has_style_effect) {
                    addresses.push(*address);
                }
            }
            for (address, left_fill) in &left.fills {
                if let Some(right_fill) = right.fills.get(address)
                    && left_fill.overlaps(right_fill)
                {
                    addresses.push(*address);
                }
                if right
                    .cells
                    .get(address)
                    .is_some_and(action_has_style_effect)
                {
                    addresses.push(*address);
                }
            }
            for address in right.fills.keys() {
                if left.cells.get(address).is_some_and(action_has_style_effect) {
                    addresses.push(*address);
                }
            }
            for (address, left_border) in &left.borders {
                if let Some(right_border) = right.borders.get(address)
                    && left_border.overlaps(right_border)
                {
                    addresses.push(*address);
                }
                if right
                    .cells
                    .get(address)
                    .is_some_and(action_has_style_effect)
                {
                    addresses.push(*address);
                }
            }
            for address in right.borders.keys() {
                if left.cells.get(address).is_some_and(action_has_style_effect) {
                    addresses.push(*address);
                }
            }
            for (address, left_protection) in &left.protections {
                if let Some(right_protection) = right.protections.get(address)
                    && left_protection.overlaps(right_protection)
                {
                    addresses.push(*address);
                }
                if right
                    .cells
                    .get(address)
                    .is_some_and(action_has_style_effect)
                {
                    addresses.push(*address);
                }
            }
            for address in right.protections.keys() {
                if left.cells.get(address).is_some_and(action_has_style_effect) {
                    addresses.push(*address);
                }
            }
            for address in left.quote_prefixes.keys() {
                if right.quote_prefixes.contains_key(address)
                    || right
                        .cells
                        .get(address)
                        .is_some_and(action_has_style_effect)
                {
                    addresses.push(*address);
                }
            }
            for address in right.quote_prefixes.keys() {
                if left.cells.get(address).is_some_and(action_has_style_effect) {
                    addresses.push(*address);
                }
            }
            for address in left.number_formats.keys() {
                if right.number_formats.contains_key(address)
                    || right
                        .cells
                        .get(address)
                        .is_some_and(action_has_style_effect)
                {
                    addresses.push(*address);
                }
            }
            for address in right.number_formats.keys() {
                if left.cells.get(address).is_some_and(action_has_style_effect) {
                    addresses.push(*address);
                }
            }
            addresses.sort_unstable();
            addresses.dedup();
            if !addresses.is_empty() {
                conflicts.push(Conflict::Cells {
                    sheet: sheet.into(),
                    position: *position,
                    addresses: addresses.into_boxed_slice(),
                });
            }

            let rows = conflicts::overlapping_keys(&left.rows, &right.rows, |left, right| {
                left.overlaps(*right)
            });
            if !rows.is_empty() {
                conflicts.push(Conflict::Rows {
                    sheet: sheet.into(),
                    position: *position,
                    rows: rows.into_boxed_slice(),
                });
            }

            let columns =
                conflicts::overlapping_keys(&left.columns, &right.columns, |left, right| {
                    left.overlaps(*right)
                });
            if !columns.is_empty() {
                conflicts.push(Conflict::Columns {
                    sheet: sheet.into(),
                    position: *position,
                    columns: columns.into_boxed_slice(),
                });
            }
        }
        ConflictSet {
            conflicts: conflicts.into_boxed_slice(),
        }
    }

    fn target_context(&self, target: Target) -> (&str, usize) {
        let projected = order::projected_position(self, target);
        match target {
            Target::Base(position) => (
                self.base
                    .inner
                    .sheets
                    .get(position)
                    .map_or("<missing sheet>", |sheet| sheet.name.as_str()),
                projected.unwrap_or(position),
            ),
            Target::Added(index) => self.added.get(index).map_or(
                ("<missing new sheet>", self.base.len().saturating_add(index)),
                |sheet| {
                    (
                        sheet.name.as_str(),
                        projected.unwrap_or_else(|| self.base.len().saturating_add(index)),
                    )
                },
            ),
        }
    }

    pub(in crate::workbook::edit) fn has_non_removal(&self) -> bool {
        self.panes.is_some()
            || self.defined_names.is_some()
            || !self.drawings.is_empty()
            || !self.svg_lifecycle.is_empty()
            || self.active.is_some()
            || self.order.as_ref().is_some_and(OrderPlan::is_effective)
            || self.sheets.values().any(|actions| !actions.is_empty())
            || !self.added.is_empty()
    }

    fn has_defined_name_scope_changes(&self) -> bool {
        !self.removed.is_empty()
            || !self.added.is_empty()
            || self.order.as_ref().is_some_and(OrderPlan::is_effective)
            || self.sheets.values().any(|actions| actions.rename.is_some())
    }

    pub(in crate::workbook::edit) fn remove_block(&self, reason: RemoveBlock, part: &str) -> Error {
        let position = self.removed.iter().next().copied().unwrap_or(0);
        let sheet = self
            .base
            .inner
            .sheets
            .get(position)
            .map_or("<missing sheet>", |sheet| sheet.name.as_str());
        Error::SheetRemoveBlocked {
            sheet: sheet.to_owned(),
            position,
            part: part.to_owned(),
            reason,
        }
    }
}

fn payload_capture_bytes(payload: &Payload) -> usize {
    match payload {
        Payload::Set(content) => content.retained_bytes(),
        Payload::SharedString { text, .. } => {
            size_of::<Payload>().saturating_add(text.as_str().len())
        },
        Payload::SharedFormula {
            reference, formula, ..
        } => size_of::<Payload>()
            .saturating_add(reference.len())
            .saturating_add(formula.as_ref().map_or(0, |formula| formula.text().len())),
        Payload::Cache(value) => size_of::<Payload>().saturating_add(value.retained_bytes()),
        Payload::Rich(text) => text.retained_bytes(),
        Payload::Clear | Payload::ClearIfPresent => size_of::<Content>(),
    }
}

fn join_normalization(error: &Error) -> JoinNormalization {
    match error {
        Error::Limit {
            kind,
            observed,
            maximum,
        } => JoinNormalization::Limit {
            kind: *kind,
            observed: *observed,
            maximum: *maximum,
        },
        Error::ResourceLimit(_) => JoinNormalization::ResourceLimit,
        Error::Package(_) => JoinNormalization::Package,
        Error::MarkupCompatibility(_) => JoinNormalization::MarkupCompatibility,
        Error::Xml(_) => JoinNormalization::Xml,
        Error::SheetNameDependency(_) | Error::Drawing(_) => JoinNormalization::Dependency,
        Error::Common(_) => JoinNormalization::Common,
        Error::InvalidLimit { .. }
        | Error::Coordinate(_)
        | Error::Range(_)
        | Error::ColumnWidth(_)
        | Error::RowHeight(_)
        | Error::DefaultHeight(_)
        | Error::DefaultWidth(_)
        | Error::Descent(_)
        | Error::Outline(_)
        | Error::SheetName(_)
        | Error::Invalid(_) => JoinNormalization::Invalid,
        Error::Allocation { .. } => JoinNormalization::Allocation,
        Error::Refused(_)
        | Error::FormControl(_)
        | Error::SheetNameConflict { .. }
        | Error::Ambiguous { .. }
        | Error::ForeignStyle
        | Error::NotWorksheet { .. }
        | Error::DanglingWebBinding { .. }
        | Error::EditBlocked { .. }
        | Error::RowEditBlocked { .. }
        | Error::ColumnEditBlocked { .. }
        | Error::DefaultsEditBlocked { .. }
        | Error::MergeEditBlocked { .. }
        | Error::TabEditBlocked { .. }
        | Error::RenameBlocked { .. }
        | Error::SheetRemoveBlocked { .. }
        | Error::CustomDataReferenced { .. } => JoinNormalization::Refused,
        Error::Signed | Error::PatchConflict { .. } | Error::DurablePatch(_) | Error::Join(_) => {
            JoinNormalization::Publication
        },
        Error::UnsupportedSelector | Error::Unsupported { .. } => JoinNormalization::Unsupported,
        // These are operation-boundary failures rather than malformed source
        // or an authored edit.  Keep the typed leaf in `JoinError`'s source;
        // the unified classifier maps the leaf to `io_error` or
        // `internal_error` respectively.
        Error::SourceChanged { .. } => JoinNormalization::Other,
        Error::Internal(_) => JoinNormalization::Other,
        #[cfg(feature = "encryption")]
        Error::Encryption(_) | Error::EncryptionPolicy { .. } => JoinNormalization::Other,
    }
}

#[cfg(test)]
mod error_classification_tests {
    use super::{Error, JoinError, JoinFailure, JoinNormalization, join_normalization};

    #[test]
    fn operation_boundary_errors_keep_their_typed_leaf_classification() {
        let source_changed = Error::SourceChanged {
            expected: litchi_core::SourceVersion::new(7, 1),
            observed: litchi_core::SourceVersion::new(7, 2),
        };
        assert_eq!(
            join_normalization(&source_changed),
            JoinNormalization::Other
        );
        assert_eq!(
            source_changed.to_string(),
            "XLSX source changed while the operation was in progress"
        );

        let internal = Error::Internal("generated staging invariant");
        assert_eq!(join_normalization(&internal), JoinNormalization::Other);
        assert_eq!(
            internal.to_string(),
            "XLSX internal invariant failed: generated staging invariant"
        );
    }

    fn boundary_join(source: Error) -> JoinError {
        let rejected = crate::Package::new()
            .expect("minimal workbook")
            .edit()
            .expect("edit");
        JoinError {
            failure: JoinFailure::Normalization(JoinNormalization::Other),
            rejected: Box::new(rejected),
            normalization_source: Some(Box::new(source)),
        }
    }

    fn assert_boundary_source(source: Error, expected: impl Fn(&Error) -> bool) {
        let join = boundary_join(source);
        let retained = join
            .normalization_source()
            .expect("normalization source retained");
        assert!(expected(retained));
        let source_chain = std::error::Error::source(&join)
            .expect("join source")
            .downcast_ref::<Error>()
            .expect("typed join source");
        assert!(expected(source_chain));

        let wrapped = Error::from(join);
        let Error::Join(join) = &wrapped else {
            panic!("JoinError converts to Error::Join");
        };
        assert!(join.normalization_source().is_some_and(&expected));
        // `Error::Join` is `#[error(transparent)]`: its standard source is the
        // `JoinError`'s own source, the retained typed leaf.
        let source_chain = std::error::Error::source(&wrapped)
            .expect("wrapped normalization source")
            .downcast_ref::<Error>()
            .expect("typed wrapped normalization source");
        assert!(expected(source_chain));
    }

    #[test]
    fn operation_boundary_join_preserves_both_typed_sources_through_wrapper() {
        let expected = litchi_core::SourceVersion::new(7, 1);
        let observed = litchi_core::SourceVersion::new(7, 2);
        assert_boundary_source(Error::SourceChanged { expected, observed }, |error| {
            matches!(error, Error::SourceChanged { .. })
        });
        assert_boundary_source(Error::Internal("generated staging invariant"), |error| {
            matches!(error, Error::Internal(_))
        });
    }
}

fn rejected_after_normalization(
    mut other: Edit,
    normalization_undo: Option<NormalizationUndo>,
) -> Box<Edit> {
    if let Some(undo) = normalization_undo {
        undo.restore(&mut other);
    }
    Box::new(other)
}

fn action_payload_matches_source(
    sheet: &Worksheet,
    stored: Option<&crate::cell::Stored>,
    address: Address,
    action: &Action,
) -> Result<bool> {
    action.payload().map_or(Ok(true), |payload| {
        payload_matches_source(sheet, stored, address, payload)
    })
}

fn payload_matches_source(
    sheet: &Worksheet,
    stored: Option<&crate::cell::Stored>,
    address: Address,
    payload: &Payload,
) -> Result<bool> {
    let source = stored.map(|value| &value.cell);
    match payload {
        Payload::Set(content) => {
            if !source.is_some_and(|value| value == &content.as_cell()) {
                return Ok(false);
            }
            let Some(stored) = stored else {
                return Ok(false);
            };
            if stored.inline_rich {
                return Ok(false);
            }
            if stored.shared_string.is_some() {
                return Ok(!sheet.has_rich_text(address)?);
            }
            Ok(true)
        },
        Payload::SharedString { index, text } => Ok(stored.is_some_and(|value| {
            value.shared_string == Some(*index)
                && matches!(&value.cell, Cell::Value(crate::Value::Text(current)) if current == text)
        })),
        Payload::Rich(text) => {
            let Some(stored) = stored else {
                return Ok(false);
            };
            // `set_rich_text` authors an inline rich string. A shared rich
            // item with identical runs is still a representation change, and
            // an unsupported inline run payload is not proof of equality.
            if !stored.inline_rich {
                return Ok(false);
            }
            Ok(sheet.rich_text(address)?.as_ref() == Some(text.as_ref()))
        },
        Payload::Cache(value) => Ok(source.is_some_and(|current| {
            matches!(
                current,
                Cell::Formula(formula)
                    if formula
                        .cached()
                        .is_some_and(|cache| cache.value() == value)
            )
        })),
        Payload::Clear => Ok(source.is_some_and(|value| matches!(value, Cell::Empty))),
        Payload::ClearIfPresent => {
            Ok(source.is_none_or(|value| matches!(value, Cell::Empty)))
        },
        Payload::SharedFormula { .. } => Ok(false),
    }
}

fn style_matches_source(style: StyleEffect, source: Option<u32>) -> bool {
    match style {
        StyleEffect::Set(style) => source == Some(style),
        StyleEffect::Reset => source.is_none(),
    }
}

fn worksheet_capture_bytes(base: &Workbook, position: usize) -> Result<usize> {
    let sheet = base
        .inner
        .sheets
        .get(position)
        .ok_or_else(|| invalid("worksheet capture target disappeared"))?;
    match base.inner.package.part_metadata(&sheet.part_uri) {
        Some(metadata) => Ok(metadata.payload_len()?),
        None => Err(invalid("worksheet capture part disappeared")),
    }
}

fn staged_cell_capture(
    actions: &SheetActions,
    address: Address,
    include_action: bool,
    include_structural: bool,
    typed: EffectiveTypedFacets,
) -> Option<CellCapture> {
    let mut capture = CellCapture::default();
    if include_action {
        if let Some(action) = actions.cells.get(&address) {
            apply_action_capture(&mut capture, action);
        }
    }
    if typed.font {
        if let Some(change) = actions.fonts.get(&address) {
            capture.apply(CaptureFacet::Font, change.retained_bytes());
        }
    }
    if typed.alignment {
        if let Some(change) = actions.alignments.get(&address) {
            capture.apply(CaptureFacet::Alignment, change.retained_bytes());
        }
    }
    if typed.fill {
        if let Some(change) = actions.fills.get(&address) {
            capture.apply(CaptureFacet::Fill, change.retained_bytes());
        }
    }
    if typed.border {
        if let Some(change) = actions.borders.get(&address) {
            capture.apply(CaptureFacet::Border, change.retained_bytes());
        }
    }
    if typed.protection {
        if let Some(change) = actions.protections.get(&address) {
            capture.apply(CaptureFacet::Protection, change.retained_bytes());
        }
    }
    if typed.quote_prefix {
        capture.apply(CaptureFacet::QuotePrefix, size_of::<Option<bool>>());
    }
    if typed.number_format {
        if let Some(value) = actions.number_formats.get(&address) {
            capture.apply(
                CaptureFacet::NumberFormat,
                size_of::<Option<Arc<str>>>()
                    .saturating_add(value.as_ref().map_or(0, |value| value.len())),
            );
        }
    }
    if include_structural {
        capture.apply(CaptureFacet::Structural, 0);
    }
    for ((row, column, _, _), action) in &actions.hyperlinks {
        if Address::at(*row, *column).ok() != Some(address) {
            continue;
        }
        capture.apply(
            CaptureFacet::Relationship,
            hyperlink_action_retained_bytes(action),
        );
    }
    capture.active.then_some(capture)
}

fn hyperlink_action_retained_bytes(action: &HyperlinkAction) -> usize {
    match action {
        HyperlinkAction::Put(value) => {
            value.reference().as_str().len()
                + value.location().map_or(0, str::len)
                + value.external_target().map_or(0, str::len)
                + value.display().map_or(0, str::len)
                + value.tooltip().map_or(0, str::len)
                + size_of::<crate::hyperlinks::Hyperlink>()
        },
        HyperlinkAction::Remove(reference) => {
            reference.as_str().len() + size_of::<crate::hyperlinks::Hyperlink>()
        },
    }
}

fn apply_action_capture(capture: &mut CellCapture, action: &Action) {
    match action {
        Action::Remove => capture.apply(CaptureFacet::Remove, size_of::<Action>()),
        Action::Update { payload, style } => {
            if let Some(payload) = payload {
                capture.apply(CaptureFacet::Payload, payload_capture_bytes(payload));
            }
            if style.is_some() {
                capture.apply(CaptureFacet::Style, size_of::<StyleEffect>());
            }
        },
    }
}

fn action_has_style_effect(action: &Action) -> bool {
    matches!(
        action,
        Action::Update { style: Some(_), .. } | Action::Remove
    )
}

fn style_staged_memory_limit(base: &Workbook) -> usize {
    let read_limit =
        usize::try_from(base.inner.package.read_limits().max_part_bytes()).unwrap_or(usize::MAX);
    read_limit.min(raw::styles_edit::StyleEditLimits::default().max_staged_bytes)
}

fn border_staged_entry_bytes() -> usize {
    size_of::<BorderChange>()
        .saturating_add(size_of::<Address>())
        .saturating_mul(BORDER_STAGED_COPY_HEADROOM)
        .saturating_add(BORDER_STAGED_NODE_BYTES)
}

fn staged_style_bytes(sheets: &BTreeMap<usize, SheetActions>) -> Result<usize> {
    sheets.values().try_fold(0usize, |retained, actions| {
        retained
            .checked_add(actions.staged_fill_bytes)
            .and_then(|retained| retained.checked_add(actions.staged_border_bytes))
            .ok_or_else(|| invalid("worksheet style retained bytes overflow"))
    })
}

fn retained_fill_bytes(sheets: &BTreeMap<usize, SheetActions>) -> Result<usize> {
    sheets.values().try_fold(0usize, |retained, actions| {
        actions.fills.values().try_fold(retained, |retained, fill| {
            retained
                .checked_add(fill.retained_bytes())
                .ok_or_else(|| invalid("worksheet fill retained bytes overflow"))
        })
    })
}

fn projected_fill_retained_bytes(
    left: &BTreeMap<usize, SheetActions>,
    right: &BTreeMap<usize, SheetActions>,
) -> Result<usize> {
    let mut retained = left.values().try_fold(0usize, |retained, actions| {
        retained
            .checked_add(actions.staged_fill_bytes)
            .ok_or_else(|| invalid("worksheet fill retained bytes overflow"))
    })?;
    for (position, actions) in right {
        for (address, incoming) in &actions.fills {
            let Some(current) = left
                .get(position)
                .and_then(|actions| actions.fills.get(address))
            else {
                retained = retained
                    .checked_add(incoming.retained_bytes())
                    .ok_or_else(|| invalid("worksheet fill retained bytes overflow"))?;
                continue;
            };
            let mut merged = current.clone();
            merged.merge(incoming);
            retained = retained
                .checked_sub(current.retained_bytes())
                .and_then(|retained| retained.checked_add(merged.retained_bytes()))
                .ok_or_else(|| invalid("worksheet fill retained bytes overflow"))?;
        }
    }
    Ok(retained)
}

fn projected_border_staged_bytes(
    left: &BTreeMap<usize, SheetActions>,
    right: &BTreeMap<usize, SheetActions>,
) -> Result<usize> {
    let mut retained = left.values().try_fold(0usize, |retained, actions| {
        retained
            .checked_add(actions.staged_border_bytes)
            .ok_or_else(|| invalid("worksheet border retained bytes overflow"))
    })?;
    for (position, actions) in right {
        for address in actions.borders.keys() {
            let already_present = left
                .get(position)
                .is_some_and(|actions| actions.borders.contains_key(address));
            if !already_present {
                retained = retained
                    .checked_add(border_staged_entry_bytes())
                    .ok_or_else(|| invalid("worksheet border retained bytes overflow"))?;
            }
        }
    }
    Ok(retained)
}

fn enforce_style_staged_memory(observed: usize, maximum: usize, scope: &str) -> Result<()> {
    if observed > maximum {
        return Err(Error::ResourceLimit(litchi_core::ResourceLimit {
            resource: litchi_core::Resource::Memory,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            limit: u64::try_from(maximum).unwrap_or(u64::MAX),
            scope: Arc::from(scope),
        }));
    }
    Ok(())
}

fn fill_join_resource_limit(observed: u64, maximum: usize) -> litchi_core::ResourceLimit {
    litchi_core::ResourceLimit {
        resource: litchi_core::Resource::Memory,
        observed,
        limit: u64::try_from(maximum).unwrap_or(u64::MAX),
        scope: Arc::from("worksheet fill staged bytes"),
    }
}

fn border_join_resource_limit(observed: u64, maximum: usize) -> litchi_core::ResourceLimit {
    litchi_core::ResourceLimit {
        resource: litchi_core::Resource::Memory,
        observed,
        limit: u64::try_from(maximum).unwrap_or(u64::MAX),
        scope: Arc::from("worksheet border staged bytes"),
    }
}

fn style_join_resource_limit(observed: u64, maximum: usize) -> litchi_core::ResourceLimit {
    litchi_core::ResourceLimit {
        resource: litchi_core::Resource::Memory,
        observed,
        limit: u64::try_from(maximum).unwrap_or(u64::MAX),
        scope: Arc::from("worksheet style staged bytes"),
    }
}

fn snapshot_join_resource_limit(observed: usize, maximum: usize) -> litchi_core::ResourceLimit {
    litchi_core::ResourceLimit {
        resource: litchi_core::Resource::Memory,
        observed: u64::try_from(observed).unwrap_or(u64::MAX),
        limit: u64::try_from(maximum).unwrap_or(u64::MAX),
        scope: Arc::from("worksheet snapshot bytes"),
    }
}

#[derive(Debug, Clone, Copy)]
struct PendingStyleRequest {
    position: usize,
    address: Address,
    source_style: u32,
    request_index: usize,
}

/// Lower all property-level style requests before any worksheet part is
/// rewritten. The stylesheet codec allocates/deduplicates the private native
/// records in one bounded pass; worksheet actions then carry only final
/// `StyleEffect::Set` values into the existing lossless worksheet writer.
fn lower_style_requests(
    base: &Workbook,
    sheets: &mut BTreeMap<usize, SheetActions>,
    parts: &mut Vec<PartChange>,
) -> Result<(bool, Option<u32>)> {
    let retained_style_bytes = retained_fill_bytes(sheets)?
        .checked_add(sheets.values().try_fold(0usize, |retained, actions| {
            retained
                .checked_add(actions.staged_border_bytes)
                .ok_or_else(|| invalid("worksheet border retained bytes overflow"))
        })?)
        .ok_or_else(|| invalid("worksheet style retained bytes overflow"))?;
    enforce_style_staged_memory(
        retained_style_bytes,
        style_staged_memory_limit(base),
        "worksheet style staged bytes",
    )?;
    let mut requests = Vec::new();
    let mut pending = Vec::new();
    let mut request_indices = HashMap::<
        (
            u32,
            FontChange,
            AlignmentChange,
            FillChange,
            BorderChange,
            ProtectionChange,
            Option<Option<bool>>,
            Option<Option<Arc<str>>>,
        ),
        usize,
    >::new();
    for (position, actions) in sheets.iter() {
        if actions.fonts.is_empty()
            && actions.alignments.is_empty()
            && actions.fills.is_empty()
            && actions.borders.is_empty()
            && actions.protections.is_empty()
            && actions.quote_prefixes.is_empty()
            && actions.number_formats.is_empty()
        {
            continue;
        }
        let data = base
            .inner
            .sheets
            .get(*position)
            .ok_or_else(|| invalid("style edit worksheet disappeared during planning"))?;
        if data.kind != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: data.name.clone(),
            });
        }
        let sheet = Worksheet {
            owner: Arc::clone(&base.inner),
            data: Arc::clone(data),
            policy: None,
        };
        let store = sheet.store()?;
        let request_count = actions
            .fonts
            .len()
            .checked_add(actions.alignments.len())
            .and_then(|count| count.checked_add(actions.fills.len()))
            .and_then(|count| count.checked_add(actions.borders.len()))
            .and_then(|count| count.checked_add(actions.protections.len()))
            .and_then(|count| count.checked_add(actions.quote_prefixes.len()))
            .and_then(|count| count.checked_add(actions.number_formats.len()))
            .ok_or_else(|| invalid("style edit request count overflows"))?;
        let mut addresses = Vec::new();
        addresses
            .try_reserve_exact(request_count)
            .map_err(|source| allocation("style edit request addresses", source))?;
        addresses.extend(actions.fonts.keys().copied());
        addresses.extend(actions.alignments.keys().copied());
        addresses.extend(actions.fills.keys().copied());
        addresses.extend(actions.borders.keys().copied());
        addresses.extend(actions.protections.keys().copied());
        addresses.extend(actions.quote_prefixes.keys().copied());
        addresses.extend(actions.number_formats.keys().copied());
        addresses.sort_unstable();
        addresses.dedup();
        requests
            .try_reserve(addresses.len())
            .map_err(|source| allocation("style edit requests", source))?;
        pending
            .try_reserve(addresses.len())
            .map_err(|source| allocation("style edit request locations", source))?;
        request_indices
            .try_reserve(addresses.len())
            .map_err(|source| allocation("style edit request index", source))?;
        for address in addresses {
            let source_style = match actions.cells.get(&address) {
                Some(Action::Update {
                    style: Some(StyleEffect::Set(style)),
                    ..
                }) => *style,
                Some(Action::Update {
                    style: Some(StyleEffect::Reset),
                    ..
                }) => 0,
                _ => store
                    .entry(address)
                    .and_then(|stored| stored.style)
                    .unwrap_or(0),
            };
            let change = actions
                .fonts
                .get(&address)
                .cloned()
                .unwrap_or_else(FontChange::new);
            let alignment = actions
                .alignments
                .get(&address)
                .cloned()
                .unwrap_or_else(AlignmentChange::new);
            let fill = actions
                .fills
                .get(&address)
                .cloned()
                .unwrap_or_else(FillChange::new);
            let border = actions
                .borders
                .get(&address)
                .cloned()
                .unwrap_or_else(BorderChange::new);
            let protection = actions
                .protections
                .get(&address)
                .cloned()
                .unwrap_or_else(ProtectionChange::new);
            let quote_prefix = actions.quote_prefixes.get(&address).copied();
            let number_format = actions.number_formats.get(&address).cloned();
            let key = (
                source_style,
                change.clone(),
                alignment,
                fill.clone(),
                border.clone(),
                protection,
                quote_prefix,
                number_format.clone(),
            );
            let request_index = if let Some(index) = request_indices.get(&key).copied() {
                index
            } else {
                let index = requests.len();
                requests.push(raw::styles_edit::StyleRequest {
                    style: source_style,
                    change: change.clone(),
                    alignment,
                    fill,
                    border,
                    protection,
                    quote_prefix,
                    number_format,
                });
                request_indices.insert(key, index);
                index
            };
            pending.push(PendingStyleRequest {
                position: *position,
                address,
                source_style,
                request_index,
            });
        }
    }
    if requests.is_empty() {
        return Ok((false, None));
    }
    let uri = base
        .inner
        .styles_uri
        .as_ref()
        .ok_or(Error::Unsupported {
            feature: "style edits require a workbook styles part",
        })?
        .clone();
    let before = base.inner.package.get_part(&uri)?.blob_arc();
    let read_limits = base.inner.package.read_limits();
    let max_part_bytes = usize::try_from(read_limits.max_part_bytes()).unwrap_or(usize::MAX);
    let limits = raw::styles_edit::StyleEditLimits {
        max_changes: MAX_STYLE_REQUESTS,
        max_fonts: MAX_STYLE_FONTS,
        max_fills: MAX_STYLE_FILLS,
        max_borders: MAX_STYLE_BORDERS,
        max_number_formats: MAX_STYLE_NUMBER_FORMATS,
        max_cell_xfs: MAX_STYLE_CELL_XFS,
        max_depth: read_limits.max_depth(),
        max_input_bytes: max_part_bytes,
        max_output_bytes: max_part_bytes,
        max_staged_bytes: style_staged_memory_limit(base),
    };
    let rewritten = raw::styles_edit::rewrite_with_limits(before.as_slice(), &requests, limits)?;
    if rewritten.styles.len() != requests.len() {
        return Err(invalid(
            "style stylesheet planning returned incomplete results",
        ));
    }
    let projected_style_count = raw::styles::parse(&rewritten.bytes)?.len();
    for pending in pending {
        let style = *rewritten
            .styles
            .get(pending.request_index)
            .ok_or_else(|| invalid("style stylesheet planning returned an invalid result index"))?;
        let actions = sheets
            .get_mut(&pending.position)
            .ok_or_else(|| invalid("style edit worksheet disappeared after planning"))?;
        actions.fonts.remove(&pending.address);
        actions.alignments.remove(&pending.address);
        actions.fills.remove(&pending.address);
        actions.borders.remove(&pending.address);
        if actions.fills.is_empty() {
            actions.staged_fill_bytes = 0;
        }
        if actions.borders.is_empty() {
            actions.staged_border_bytes = 0;
        }
        actions.protections.remove(&pending.address);
        actions.quote_prefixes.remove(&pending.address);
        actions.number_formats.remove(&pending.address);
        if style == pending.source_style {
            continue;
        }
        match actions.cells.entry(pending.address) {
            Entry::Vacant(entry) => {
                entry.insert(Action::style(style));
            },
            Entry::Occupied(mut entry) => {
                entry.get_mut().set_style(StyleEffect::Set(style));
            },
        }
    }
    if rewritten.bytes.as_slice() == before.as_slice() {
        return Ok((false, Some(projected_style_count)));
    }
    parts.push(PartChange {
        uri,
        before,
        after: Arc::new(rewritten.bytes),
        before_source: None,
        after_source: None,
    });
    Ok((true, Some(projected_style_count)))
}

fn validate_store_style_count(store: &crate::cell::Store, len: u32) -> Result<()> {
    if let Some(entry) = store
        .entries()
        .iter()
        .find(|entry| entry.style.is_some_and(|style| style >= len))
    {
        return Err(invalid(format!(
            "worksheet cell {} references shared style {}, but the workbook contains {len} cell formats",
            entry.address,
            entry.style.unwrap_or_default()
        )));
    }
    if let Some(entry) = store
        .row_entries()
        .iter()
        .find(|entry| entry.properties.style.is_some_and(|style| style >= len))
    {
        return Err(invalid(format!(
            "worksheet row {} references shared style {}, but the workbook contains {len} cell formats",
            entry.index,
            entry.properties.style.unwrap_or_default()
        )));
    }
    if let Some(entry) = store
        .column_entries()
        .iter()
        .find(|entry| entry.properties.style.is_some_and(|style| style >= len))
    {
        return Err(invalid(format!(
            "worksheet column {} references shared style {}, but the workbook contains {len} cell formats",
            entry.first,
            entry.properties.style.unwrap_or_default()
        )));
    }
    Ok(())
}

struct HyperlinkRewrite {
    values: Vec<crate::hyperlinks::Hyperlink>,
    relationship_ids: Vec<Option<String>>,
    verification_relationships: Relationships,
}

#[allow(clippy::too_many_arguments)]
fn project_hyperlink_actions(
    sheet: &str,
    owner: &litchi_opc::PackURI,
    xml: &[u8],
    source_relationships: &Relationships,
    actions: BTreeMap<(u32, u32, u32, u32), HyperlinkAction>,
    changes: &mut Vec<Change>,
    relationship_changes: &mut Vec<RelationshipChange>,
) -> Result<Option<HyperlinkRewrite>> {
    if actions.is_empty() {
        return Ok(None);
    }
    let mut staged =
        crate::hyperlinks::codec::parse_with_relationship_ids(xml, source_relationships)?;
    let relationship_type = crate::hyperlinks::codec::relationship_type(xml)?;
    let mut used_relationship_ids = HashSet::new();
    used_relationship_ids
        .try_reserve(source_relationships.len().saturating_add(actions.len()))
        .map_err(|source| allocation("worksheet hyperlink relationship-ID index", source))?;
    used_relationship_ids.extend(
        source_relationships
            .iter()
            .map(|relationship| relationship.r_id().to_owned()),
    );
    let mut next_relationship_id = 1_u32;
    let mut semantic_changes = Vec::new();
    semantic_changes
        .try_reserve(actions.len())
        .map_err(|source| allocation("worksheet hyperlink semantic changes", source))?;
    let mut physical_changes = Vec::new();
    physical_changes
        .try_reserve(actions.len())
        .map_err(|source| allocation("worksheet hyperlink relationship changes", source))?;
    let mut verification_relationships = source_relationships.clone();

    for (key, action) in actions {
        let existing_index = staged.iter().position(|candidate| {
            crate::hyperlinks::model::reference_key(candidate.value.reference().range()) == key
        });
        let before = existing_index.map(|index| staged[index].value.clone());
        let previous_relationship_id =
            existing_index.and_then(|index| staged[index].relationship_id.clone());
        let (reference, after) = match action {
            HyperlinkAction::Put(value) => (value.reference().clone(), Some(value)),
            HyperlinkAction::Remove(reference) => (reference, None),
        };
        if before == after {
            continue;
        }

        let next_relationship = match after.as_ref().and_then(|value| value.external_target()) {
            Some(target) => {
                let relationship_id = previous_relationship_id.clone().map_or_else(
                    || {
                        allocate_hyperlink_relationship_id(
                            &mut used_relationship_ids,
                            &mut next_relationship_id,
                        )
                    },
                    Ok,
                )?;
                Some(Relationship::new_with_mode(
                    relationship_id,
                    relationship_type.to_owned(),
                    target.to_owned(),
                    owner.base_uri().to_owned(),
                    TargetMode::External,
                ))
            },
            None => None,
        };
        let previous_relationship = previous_relationship_id
            .as_deref()
            .and_then(|relationship_id| source_relationships.get(relationship_id))
            .cloned();
        let staged_relationship_id = next_relationship
            .as_ref()
            .map(|relationship| relationship.r_id().to_owned());
        if !same_hyperlink_relationship(previous_relationship.as_ref(), next_relationship.as_ref())
        {
            let delta = RelationshipChange {
                owner: owner.clone(),
                before: previous_relationship,
                after: next_relationship,
                before_source: None,
                after_source: None,
            };
            apply_relationship_change_to_collection(&mut verification_relationships, &delta)?;
            physical_changes.push(delta);
        }

        match (existing_index, after.clone()) {
            (Some(index), Some(value)) => {
                staged[index] = crate::hyperlinks::codec::ParsedHyperlink {
                    value,
                    relationship_id: staged_relationship_id,
                };
            },
            (Some(index), None) => {
                staged.remove(index);
            },
            (None, Some(value)) => staged.push(crate::hyperlinks::codec::ParsedHyperlink {
                value,
                relationship_id: staged_relationship_id,
            }),
            (None, None) => {},
        }
        semantic_changes.push(Change::Hyperlink {
            sheet: sheet.into(),
            reference,
            before,
            after,
        });
    }
    if semantic_changes.is_empty() {
        return Ok(None);
    }
    changes.extend(semantic_changes);
    relationship_changes.extend(physical_changes);
    Ok(Some(HyperlinkRewrite {
        values: staged.iter().map(|value| value.value.clone()).collect(),
        relationship_ids: staged
            .into_iter()
            .map(|value| value.relationship_id)
            .collect(),
        verification_relationships,
    }))
}

fn allocate_hyperlink_relationship_id(
    used: &mut HashSet<String>,
    next: &mut u32,
) -> Result<String> {
    loop {
        let candidate = format!("rId{next}");
        *next = (*next)
            .checked_add(1)
            .ok_or_else(|| invalid("worksheet hyperlink relationship-ID space is exhausted"))?;
        if used.insert(candidate.clone()) {
            return Ok(candidate);
        }
    }
}

fn same_hyperlink_relationship(left: Option<&Relationship>, right: Option<&Relationship>) -> bool {
    match (left, right) {
        (None, None) => true,
        (Some(left), Some(right)) => {
            left.r_id() == right.r_id()
                && left.reltype() == right.reltype()
                && left.target_ref() == right.target_ref()
                && left.target_mode() == right.target_mode()
        },
        (None, Some(_)) | (Some(_), None) => false,
    }
}

fn apply_relationship_change_to_collection(
    relationships: &mut Relationships,
    change: &RelationshipChange,
) -> Result<()> {
    let relationship_id = change
        .after
        .as_ref()
        .or(change.before.as_ref())
        .map(Relationship::r_id)
        .ok_or_else(|| invalid("worksheet hyperlink relationship delta has no identity"))?
        .to_owned();
    relationships.remove(&relationship_id);
    if let Some(after) = &change.after {
        relationships.try_add_relationship(
            after.reltype().to_owned(),
            after.target_ref().to_owned(),
            relationship_id,
            after.target_mode(),
        )?;
    }
    Ok(())
}

fn take_effective_renames(
    base: &Workbook,
    sheets: &mut BTreeMap<usize, SheetActions>,
) -> Result<Vec<(usize, Name)>> {
    let mut effective = Vec::new();
    for (position, actions) in sheets {
        let Some(name) = actions.rename.take() else {
            continue;
        };
        let data = base
            .inner
            .sheets
            .get(*position)
            .ok_or_else(|| invalid(format!("renamed sheet position {position} disappeared")))?;
        if name.as_str() != data.name {
            effective.push((*position, name));
        }
    }
    Ok(effective)
}

fn validate_final_sheet_names<'a>(
    base: &'a Workbook,
    final_order: &FinalOrder,
    added: &'a [Added],
    rename_by_position: &HashMap<usize, &'a Name>,
) -> Result<()> {
    let final_len = base
        .inner
        .sheets
        .len()
        .checked_add(added.len())
        .ok_or_else(|| invalid("final worksheet count overflow"))?;
    if final_len > raw::catalog_edit::MAX_SHEETS {
        let first = added
            .first()
            .ok_or_else(|| invalid("worksheet limit exceeded without a creation"))?;
        return Err(Error::TabEditBlocked {
            sheet: first.name.as_str().to_owned(),
            position: base.inner.sheets.len(),
            reason: TabEditBlock::SheetLimit,
        });
    }
    let mut final_names = HashMap::<&str, usize>::new();
    final_names
        .try_reserve(final_len)
        .map_err(|source| allocation("final sheet-name index", source))?;
    for (position, target) in final_order.targets.iter().copied().enumerate() {
        let (name, key) = match target {
            Target::Base(identity) => {
                let data =
                    base.inner.sheets.get(identity).ok_or_else(|| {
                        invalid("existing worksheet disappeared from final names")
                    })?;
                rename_by_position
                    .get(&identity)
                    .map_or((data.name.as_str(), data.name_key.as_ref()), |name| {
                        (name.as_str(), name.identity_key())
                    })
            },
            Target::Added(index) => {
                let created = added
                    .get(index)
                    .ok_or_else(|| invalid("created worksheet disappeared from final names"))?;
                (created.name.as_str(), created.name.identity_key())
            },
        };
        if let Some(first) = final_names.insert(key, position) {
            return Err(Error::SheetNameConflict {
                name: name.to_owned(),
                first,
                second: position,
            });
        }
    }
    Ok(())
}

fn ensure_defined_name_edit_is_composable(
    requested_defined_names: Option<&[raw::DefinedName]>,
    requested_order: Option<&OrderPlan>,
    sheets: &BTreeMap<usize, SheetActions>,
    added: &[Added],
) -> Result<()> {
    if requested_defined_names.is_some()
        && (requested_order.is_some_and(OrderPlan::is_effective)
            || !added.is_empty()
            || sheets.values().any(|actions| actions.rename.is_some()))
    {
        return Err(Error::Unsupported {
            feature: "combining defined-name replacement with structural sheet edits",
        });
    }
    Ok(())
}

fn validate_web_edit(
    base: &Workbook,
    requested_panes: Option<&PanesAction>,
    sheets: &BTreeMap<usize, SheetActions>,
    added: &[Added],
) -> Result<()> {
    let validates_web = requested_panes.is_some()
        || sheets.values().any(|actions| actions.web.is_some())
        || added.iter().any(|sheet| sheet.actions.web.is_some());
    if !validates_web {
        return Ok(());
    }
    let final_panes = match requested_panes {
        Some(PanesAction::Put { panes, .. }) => Some(panes),
        Some(PanesAction::Remove) => None,
        None => base.task_panes()?,
    };
    let refs = match final_panes {
        Some(panes) => crate::web::Refs::from_panes(panes)?,
        None => crate::web::Refs::new(std::iter::empty::<&str>())?,
    };
    for (position, data) in base.inner.sheets.iter().enumerate() {
        if data.kind != WorksheetKind::Worksheet {
            continue;
        }
        if let Some(bindings) = sheets
            .get(&position)
            .and_then(|actions| actions.web.as_ref())
        {
            codec::check_web_bindings(&refs, &data.name, bindings)?;
        } else {
            let sheet = Worksheet {
                owner: Arc::clone(&base.inner),
                data: Arc::clone(data),
                policy: None,
            };
            codec::check_web_bindings(&refs, &data.name, sheet.web_bindings()?)?;
        }
    }
    for sheet in added {
        if let Some(bindings) = &sheet.actions.web {
            codec::check_web_bindings(&refs, sheet.name.as_str(), bindings)?;
        }
    }
    Ok(())
}

fn transfer_action(state: State, source: Address, target: Address) -> Result<Action> {
    let State::Cell {
        content,
        style,
        shared_string,
    } = state
    else {
        return Ok(Action::Remove);
    };
    let payload = match (content, shared_string) {
        (Cell::Empty, _) => Payload::Clear,
        (Cell::Value(crate::Value::Text(text)), Some(shared_string)) => Payload::SharedString {
            index: shared_string.raw(),
            text,
        },
        (Cell::Value(value), _) => {
            let content = Content::Value(value);
            content.validate_for_write()?;
            Payload::Set(content)
        },
        (Cell::Formula(formula), _) => {
            if !matches!(formula.kind(), FormulaKind::Scalar) {
                return Err(Error::Unsupported {
                    feature: "copying range-owned formulas",
                });
            }
            let translated = crate::formula::shared::translate_formula(
                formula.text(),
                source.row().get() + 1,
                source.column().get() + 1,
                target.row().get() + 1,
                target.column().get() + 1,
            );
            Payload::Set(Content::Formula(Formula::new(translated)?))
        },
        (Cell::Unknown(_), _) => {
            return Err(Error::Unsupported {
                feature: "copying unmodeled cell encodings",
            });
        },
    };
    let style = Some(match style {
        crate::StyleState::Default => StyleEffect::Reset,
        crate::StyleState::Shared(key) => StyleEffect::Set(key.raw()),
    });
    Ok(Action::Update {
        payload: Some(payload),
        style,
    })
}

fn transfer_rich_action(rich: crate::RichText, state: &State) -> Action {
    let style = match state {
        State::Cell {
            style: crate::StyleState::Default,
            ..
        }
        | State::Missing => Some(StyleEffect::Reset),
        State::Cell {
            style: crate::StyleState::Shared(key),
            ..
        } => Some(StyleEffect::Set(key.raw())),
    };
    Action::Update {
        payload: Some(Payload::Rich(Box::new(rich))),
        style,
    }
}

const fn rectangles_overlap(left: Rect, right: Rect) -> bool {
    left.start().row().get() < right.end().0
        && right.start().row().get() < left.end().0
        && left.start().column().get() < right.end().1
        && right.start().column().get() < left.end().1
}
