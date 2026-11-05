//! Unified ODS writes over the leaf `Package::edit` transaction (record
//! 0277).
//!
//! Every verb resolves its targets against this immutable snapshot, admits
//! the unified property shape through the registry, and stages the whole
//! operation in one leaf [`litchi_ods::document::Edit`]: cell values,
//! formulas, styles and merge spans through the worksheet family (one
//! `set_cells` per sheet), sheets and rows through the leaf's provenance
//! splices, named ranges through the definitions family, annotations through
//! the annotation family and `meta.xml` through `Edit::set_metadata`. The
//! leaf commit publishes a new package; this handle never changes.

pub(crate) mod coalesce;
pub(crate) mod find;
pub(crate) mod import;

use std::collections::BTreeMap;
use std::num::NonZeroUsize;

use litchi_core::selector::{Path, Predicate, Step};
use litchi_core::{FileFormat, Position};
use litchi_ods::annotations::Annotation;
use litchi_ods::document::{Edit, LogicalRowOperation, SheetPosition};
use litchi_ods::model::names::{Definition, Range as NamedRange, Scope};
use litchi_ods::worksheet::{CellChange, Selector};
use litchi_ods::{Cell, CellValue, Merge, Package, Row, Sheet};

use super::node::{Item, Node};
use super::spec::{
    CELL_WRITE_TABLE, COMMENT_WRITE_TABLE, CellSetter, CommentSetter, Kind, MERGE_WRITE_TABLE,
    MergeSetter, NAME_WRITE_TABLE, NameSetter, SHEET_WRITE_TABLE, SheetSetter,
};
use super::{Nav, Ods, core_error, odf_error};
use crate::error::{Code, Error, Result};
use crate::new_node::NewNode;
use crate::path::PathArg;
use crate::placement::Placement;
use crate::props::{IntoProps, PropValue, Props};
use crate::registry::Operation;
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use crate::value::Scalar;

/// The native proof of one unified ODS write.
pub(crate) enum Patch {
    /// A primary document transaction.
    Document(litchi_ods::document::Patch),
    /// A `/metadata` commit of the metadata family (record 0288).
    Metadata(litchi_ods::metadata::Patch),
}

impl Patch {
    pub(crate) fn is_noop(&self) -> bool {
        match self {
            Self::Document(patch) => patch.is_noop(),
            Self::Metadata(patch) => patch.is_noop(),
        }
    }
}

/// The native proof and immutable endpoints of one unified ODS write.
pub(crate) struct Commit {
    pub(crate) before: Ods,
    pub(crate) after: Ods,
    pub(crate) patch: Patch,
    pub(crate) operations: usize,
    pub(crate) snapshot_bytes: usize,
}

/// The bytes one staged cell change charges beyond its text.
const CELL_CHARGE: usize = 64;

/// One resolved target, owned so the write does not borrow a node.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Target {
    Workbook,
    Sheet(u32),
    Cell {
        sheet: u32,
        row: u32,
        column: u32,
    },
    Range {
        sheet: u32,
        rect: (u32, u32, u32, u32),
    },
    Row {
        sheet: u32,
        index: u32,
    },
    Column {
        sheet: u32,
        index: u32,
    },
    Name(usize),
    Comment(usize),
    Merge {
        sheet: u32,
        rect: (u32, u32, u32, u32),
    },
    /// A worksheet feature, written through `feature_write` (record 0327).
    Feature(u32),
}

impl Target {
    fn of(node: &Node<'_>) -> Self {
        let sheet = node.sheet_position();
        match node.item() {
            Item::Workbook => Self::Workbook,
            Item::Sheet(_) => Self::Sheet(sheet),
            Item::Cell(cell) => Self::Cell {
                sheet,
                row: cell.at.0,
                column: cell.at.1,
            },
            Item::Range(range) => Self::Range {
                sheet,
                rect: range.rect,
            },
            Item::Row(row) => Self::Row {
                sheet,
                index: row.index,
            },
            Item::Column(column) => Self::Column {
                sheet,
                index: column.index,
            },
            Item::NamedRange(index) => Self::Name(*index),
            Item::Comment(comment) => Self::Comment(comment.entry),
            Item::Merge(merge) => Self::Merge {
                sheet,
                rect: merge.rect,
            },
            Item::Feature(_) => Self::Feature(sheet),
        }
    }

    fn element(self) -> Kind {
        match self {
            Self::Workbook => Kind::Workbook,
            Self::Sheet(_) => Kind::Sheet,
            Self::Cell { .. } => Kind::Cell,
            Self::Range { .. } => Kind::Range,
            Self::Row { .. } => Kind::Row,
            Self::Column { .. } => Kind::Column,
            Self::Name(_) => Kind::NamedRange,
            Self::Comment(_) => Kind::Comment,
            Self::Merge { .. } => Kind::MergeCell,
            Self::Feature(_) => Kind::Hyperlink,
        }
    }
}

fn refusal(code: Code, message: &'static str) -> Error {
    Error::new(code, message).with_format(FileFormat::Ods)
}

fn limit(kind: LimitKind, observed: usize, maximum: usize) -> Error {
    Error::from(TransactionError::Limit {
        kind,
        observed,
        maximum,
    })
    .with_format(FileFormat::Ods)
}

/// A leaf staging failure. The leaf reports a document shape it will not
/// rewrite (dependencies, unmodeled markup) and a malformed authored value
/// with the same `InvalidFormat` it uses for corrupt input; at a write call
/// site the input already parsed, so those authored failures are re-coded as
/// `code` while keeping the typed leaf error. A typed `Refused` keeps its
/// `RefusalKind` classification.
pub(super) fn leaf(error: litchi_core::Error, code: Code) -> Error {
    match error {
        litchi_core::Error::InvalidFormat(_)
        | litchi_core::Error::ParseError(_)
        | litchi_core::Error::Unsupported(_) => core_error(error).caused_by_argument(code),
        other => core_error(other),
    }
}

/// The same caller-argument classification for the ODF-owned leaf errors.
/// Typed limits, refusals, conflicts and source publication diagnostics retain
/// their native classification and payload.
pub(super) fn odf_leaf(error: litchi_odf_common::Error, code: Code) -> Error {
    match error {
        litchi_odf_common::Error::Core(
            error @ (litchi_core::Error::InvalidFormat(_)
            | litchi_core::Error::ParseError(_)
            | litchi_core::Error::Unsupported(_)),
        ) => core_error(error).caused_by_argument(code),
        other => odf_error(other),
    }
}

impl Ods {
    /// Admit an edit: the source is current. A flat spreadsheet edits its
    /// packaged projection (record 0292).
    pub(crate) fn admit_edit(&self) -> Result<()> {
        self.check_source()
    }

    /// The package one edit starts from.
    pub(crate) fn editable_package(&self) -> Result<Package> {
        self.admit_edit()?;
        self.materialized()
    }

    /// A new leaf edit over this snapshot.
    pub(super) fn leaf_edit(&self) -> Result<Edit> {
        self.editable_package()?.edit().map_err(odf_error)
    }

    /// A new handle over a committed package, keeping this handle's origin.
    /// A flat spreadsheet republishes the edit of its projection into its
    /// flat bytes, element by element (record 0292).
    pub(crate) fn with_package(&self, package: Package) -> Result<Self> {
        if let Some(flat) = self.flat_bytes() {
            let options = self
                .flat_options()
                .expect("flat bytes always come from a flat owner");
            let flat_limits = litchi_odf_common::flat::Limits::from_options(options);
            let parts =
                litchi_odf_common::flat::Parts::split(flat, &flat_limits).map_err(odf_error)?;
            let bytes = parts
                .merge_package_projection(flat, package.as_bytes(), &flat_limits)
                .map_err(odf_error)?;
            let next = Self::from_flat(bytes, &crate::OpenOptions::new().odf_options(options))?;
            return Ok(next.with_origin(self.origin()));
        }
        Self::from_package_with_options(package, self.origin(), self.odf_options())
    }

    /// Whether two handles share one immutable state.
    pub(crate) fn same_snapshot(&self, other: &Self) -> bool {
        self.same_inner(other)
    }

    /// Whether two handles publish the same exact bytes.
    pub(crate) fn same_content(&self, other: &Self) -> Result<bool> {
        if self.same_inner(other) {
            return Ok(true);
        }
        if let (Some(left), Some(right)) = (self.flat_bytes(), other.flat_bytes()) {
            return Ok(left == right);
        }
        Ok(self.materialized()?.as_bytes() == other.materialized()?.as_bytes())
    }

    /// Set properties on a path or a scoped query under the default
    /// semantic limits.
    pub(crate) fn set<'p, 'v>(
        &self,
        target: TargetArg<'p>,
        props: impl IntoProps<'v>,
    ) -> Result<Commit> {
        self.set_with_limits(target, props, &SemanticLimits::default())
    }

    /// Set cell, range, sheet, named-range, annotation or `/metadata`
    /// properties in one leaf edit.
    pub(crate) fn set_with_limits<'p, 'v>(
        &self,
        target: TargetArg<'p>,
        props: impl IntoProps<'v>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        self.admit_edit()?;
        let props = Props::from_props(props)?;
        if props.len() > limits.max_operations() {
            return Err(limit(
                LimitKind::Operations,
                props.len(),
                limits.max_operations(),
            ));
        }
        if let Some(commit) = super::feature_write::try_set(self, &target, &props, limits)? {
            self.check_source()?;
            return Ok(commit);
        }
        let result = if super::metadata::is_target(&target) && !is_sheet_property(&props) {
            super::metadata::set(self, &props, limits)
        } else {
            let targets = resolve(self, &target, limits)?;
            set_targets(self, &targets, &props, limits)
        };
        self.check_source()?;
        result
    }

    /// Remove under the default semantic limits.
    pub(crate) fn remove<'p>(&self, target: TargetArg<'p>) -> Result<Commit> {
        self.remove_with_limits(target, &SemanticLimits::default())
    }

    /// Remove cell content, sheets, rows, named ranges, annotations or merge
    /// spans in one leaf edit.
    pub(crate) fn remove_with_limits<'p>(
        &self,
        target: TargetArg<'p>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        self.admit_edit()?;
        if let Some(commit) = super::feature_write::try_remove(self, &target, limits)? {
            self.check_source()?;
            return Ok(commit);
        }
        let targets = resolve(self, &target, limits)?;
        let result = remove_targets(self, &targets, limits);
        self.check_source()?;
        result
    }

    /// Add a worksheet, a row, a named range or an annotation.
    pub(crate) fn add<'p, 'n>(
        &self,
        parent: PathArg<'p>,
        node: NewNode<'n>,
    ) -> Result<(Commit, Path<'static>)> {
        self.add_with_limits(parent, node, &SemanticLimits::default())
    }

    /// [`Self::add`] under explicit limits.
    pub(crate) fn add_with_limits<'p, 'n>(
        &self,
        parent: PathArg<'p>,
        node: NewNode<'n>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        self.admit_edit()?;
        let parent = Target::of(&crate::path::resolve::get(&Nav::new(self), &parent)?);
        let result = add(self, parent, &node, limits);
        self.check_source()?;
        result
    }

    /// Move a worksheet or a row.
    pub(crate) fn move_to<'p, 'q>(
        &self,
        path: PathArg<'p>,
        placement: Placement<'q>,
    ) -> Result<Commit> {
        self.move_to_with_limits(path, placement, &SemanticLimits::default())
    }

    /// [`Self::move_to`] under explicit limits.
    pub(crate) fn move_to_with_limits<'p, 'q>(
        &self,
        path: PathArg<'p>,
        placement: Placement<'q>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        self.admit_edit()?;
        let target = Target::of(&crate::path::resolve::get(&Nav::new(self), &path)?);
        let result = move_target(self, target, &placement, limits);
        self.check_source()?;
        result
    }

    /// Swap two worksheets or two rows.
    pub(crate) fn swap<'p, 'q>(&self, left: PathArg<'p>, right: PathArg<'q>) -> Result<Commit> {
        self.swap_with_limits(left, right, &SemanticLimits::default())
    }

    /// [`Self::swap`] under explicit limits.
    pub(crate) fn swap_with_limits<'p, 'q>(
        &self,
        left: PathArg<'p>,
        right: PathArg<'q>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        self.admit_edit()?;
        let nav = Nav::new(self);
        let left = Target::of(&crate::path::resolve::get(&nav, &left)?);
        let right = Target::of(&crate::path::resolve::get(&nav, &right)?);
        let result = swap_targets(self, left, right, limits);
        self.check_source()?;
        result
    }
}

fn is_sheet_property(props: &Props<'_>) -> bool {
    props
        .iter()
        .any(|(key, _)| key.eq_ignore_ascii_case("name"))
}

/// Resolve a path or a scoped query into owned targets.
fn resolve(file: &Ods, target: &TargetArg<'_>, limits: &SemanticLimits) -> Result<Vec<Target>> {
    let mut targets = Vec::new();
    if let Some(path) = target.path() {
        targets.push(Target::of(&crate::path::resolve::get(
            &Nav::new(file),
            path,
        )?));
    } else if let Some(query) = target.query() {
        let walk = crate::query::run(Nav::new(file), query)?;
        if !walk.plan().diagnostics().is_empty() {
            return Err(refusal(
                Code::InvalidSelector,
                "a mutation query must not carry compile diagnostics",
            ));
        }
        for node in walk {
            if targets.len() >= limits.max_expanded_targets() {
                return Err(limit(
                    LimitKind::ExpandedTargets,
                    targets.len().saturating_add(1),
                    limits.max_expanded_targets(),
                ));
            }
            targets.push(Target::of(&node?));
        }
    } else {
        return Err(refusal(Code::InvalidArgument, "the write target is empty"));
    }
    Ok(targets)
}

/// Commit `edit` and wrap its endpoints.
pub(super) fn finish(
    file: &Ods,
    edit: Edit,
    operations: usize,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let commit = edit.commit().map_err(odf_error)?;
    let changed = commit.diagnostics().changed();
    let (package, patch) = commit.into_parts();
    let before = file.materialized()?.as_bytes().len();
    let snapshot_bytes = if changed {
        before.saturating_add(package.as_bytes().len())
    } else {
        0
    };
    if snapshot_bytes > limits.max_snapshot_bytes() {
        return Err(limit(
            LimitKind::SnapshotBytes,
            snapshot_bytes,
            limits.max_snapshot_bytes(),
        ));
    }
    let after = if changed {
        file.with_package(package)?
    } else {
        file.clone()
    };
    Ok(Commit {
        before: file.clone(),
        after,
        patch: Patch::Document(patch),
        operations,
        snapshot_bytes,
    })
}

/// The exact no-op commit of this snapshot.
fn no_op(file: &Ods, limits: &SemanticLimits) -> Result<Commit> {
    finish(file, file.leaf_edit()?, 0, limits)
}

// ---------------------------------------------------------------- set

fn set_targets(
    file: &Ods,
    targets: &[Target],
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let Some(first) = targets.first() else {
        return no_op(file, limits);
    };
    match first.element() {
        Kind::Cell | Kind::Range => set_cells(file, targets, props, limits),
        Kind::Sheet => set_sheet(file, targets, props, limits),
        Kind::NamedRange => set_name(file, targets, props, limits),
        Kind::Comment => set_comment(file, targets, props, limits),
        _ => Err(refusal(
            Code::UnsupportedElement,
            "ODS writes set cells, ranges, sheets, named ranges and annotations",
        )
        .with_suggestion("see `capabilities()` for the settable ODS elements")),
    }
}

/// The cell assignments of one `set` call, in key order.
#[derive(Debug, Clone, Default)]
pub(super) struct CellAssignments {
    value: Option<Option<(CellValue, String)>>,
    clear: bool,
    formula: Option<Option<String>>,
    style: Option<Option<String>>,
    merge: Option<bool>,
    bytes: usize,
}

impl CellAssignments {
    /// Admit `props` through the cell registry (and the merge registry for
    /// ranges).
    pub(super) fn admit(props: &Props<'_>, ranges: bool) -> Result<Self> {
        let mut out = Self::default();
        if props.is_empty() {
            return Err(refusal(
                Code::InvalidArgument,
                "a cell set needs at least one property",
            ));
        }
        for (key, value) in props.iter() {
            let name = key.as_str();
            if ranges && name.eq_ignore_ascii_case("merge") {
                let (_, entry) = MERGE_WRITE_TABLE.resolve(name)?;
                let entry = *entry;
                entry
                    .prepare(Operation::Set, value, |MergeSetter::Merge, value| {
                        out.merge = Some(matches!(value, Some(Scalar::Bool(true))));
                        Ok(())
                    })
                    .map_err(|error| error.with_format(FileFormat::Ods))?;
                continue;
            }
            let (_, entry) = CELL_WRITE_TABLE.resolve(name).map_err(|error| {
                error
                    .with_format(FileFormat::Ods)
                    .with_property(name.to_owned())
            })?;
            let entry = *entry;
            entry
                .prepare(Operation::Set, value, |setter, scalar| {
                    out.bytes = out.bytes.saturating_add(value_bytes(value));
                    match setter {
                        CellSetter::Value => {
                            out.value = Some(scalar.map(cell_value).transpose()?);
                        },
                        CellSetter::Clear => {
                            if !matches!(scalar, Some(Scalar::Bool(true))) {
                                return Err(refusal(
                                    Code::InvalidValue,
                                    "`clear` accepts only true",
                                )
                                .with_property("clear"));
                            }
                            out.clear = true;
                        },
                        CellSetter::Formula => {
                            out.formula = Some(match scalar {
                                Some(Scalar::Text(text)) => Some(formula_token(text)),
                                None => None,
                                Some(_) => {
                                    return Err(refusal(
                                        Code::InvalidValue,
                                        "an ODS formula requires text",
                                    )
                                    .with_property("formula"));
                                },
                            });
                        },
                        CellSetter::Style => {
                            out.style = Some(match scalar {
                                Some(Scalar::Text(text)) => Some(text.to_owned()),
                                None => None,
                                Some(_) => {
                                    return Err(refusal(
                                        Code::InvalidValue,
                                        "an ODS cell style requires a style name",
                                    )
                                    .with_property("style"));
                                },
                            });
                        },
                    }
                    Ok(())
                })
                .map_err(|error| error.with_format(FileFormat::Ods))?;
        }
        if out.clear && (out.value.is_some() || out.formula.is_some()) {
            return Err(refusal(
                Code::InvalidArgument,
                "`clear` cannot be combined with `value` or `formula`",
            ));
        }
        Ok(out)
    }

    /// Whether only content keys (no merge, no style) are assigned.
    pub(super) fn is_content_only(&self) -> bool {
        self.merge.is_none() && self.style.is_none()
    }

    /// The staged bytes of one target.
    pub(super) fn charge(&self) -> usize {
        CELL_CHARGE.saturating_add(self.bytes)
    }

    /// The cell `existing` becomes; `None` when nothing changes.
    pub(super) fn apply(&self, existing: Option<&Cell>) -> Result<Option<Cell>> {
        let mut cell = existing.cloned().unwrap_or_else(Cell::empty);
        cell.repeat = NonZeroUsize::MIN;
        if self.clear {
            clear_content(&mut cell);
        }
        if let Some(value) = &self.value {
            let (new_value, new_text) = match value {
                Some((value, text)) => (value.clone(), text.clone()),
                None => (CellValue::Empty, String::new()),
            };
            // The cell's hyperlinks live in its text (`text:a`), and the
            // registry has no writable link: a value set that changes the
            // text would drop them, so it is refused until `clear` removes
            // them explicitly; the same value keeps them (record 0304).
            let same = cell.value == new_value && cell.text == new_text;
            if !cell.hyperlinks.is_empty() && !same {
                return Err(Error::new(
                    Code::UnsupportedCapability,
                    "the ODS cell holds a hyperlink, which a value set would drop",
                )
                .with_suggestion("clear the cell first (`clear=true`), then set its value"));
            }
            cell.value = new_value;
            cell.text = new_text;
            cell.formula = None;
        }
        if let Some(formula) = &self.formula {
            match formula {
                Some(formula) => cell
                    .set_formula(formula.clone())
                    .map_err(|error| leaf(error, Code::InvalidValue).with_property("formula"))?,
                None => cell.clear_formula(),
            }
        }
        if let Some(style) = &self.style {
            match style {
                Some(style) => cell
                    .set_style_name(style.clone())
                    .map_err(|error| leaf(error, Code::InvalidValue).with_property("style"))?,
                None => cell.clear_style_name(),
            }
        }
        let unchanged = existing.is_some_and(|existing| {
            let mut same = existing.clone();
            same.repeat = NonZeroUsize::MIN;
            same == cell
        }) || (existing.is_none() && cell == Cell::empty());
        Ok((!unchanged).then_some(cell))
    }
}

/// Remove a cell's value, text, formula, links and external range source,
/// keeping its style and merge role.
fn clear_content(cell: &mut Cell) {
    cell.value = CellValue::Empty;
    cell.text.clear();
    cell.formula = None;
    cell.hyperlinks.clear();
    cell.range_source = None;
}

fn value_bytes(value: &PropValue<'_>) -> usize {
    value.as_text().map_or(16, str::len)
}

/// The ODF value and displayed text of a scalar.
fn cell_value(scalar: Scalar<'_>) -> Result<(CellValue, String)> {
    Ok(match scalar {
        Scalar::Bool(value) => (
            CellValue::Boolean(value),
            if value { "TRUE" } else { "FALSE" }.to_owned(),
        ),
        Scalar::Int(value) => {
            #[allow(
                clippy::cast_precision_loss,
                reason = "ODF stores every number as office:value; a wider integer keeps its text"
            )]
            let number = value as f64;
            (CellValue::Number(number), value.to_string())
        },
        Scalar::Number(value) => (CellValue::Number(value), value.to_string()),
        Scalar::Text(text) | Scalar::Token(text) => {
            (CellValue::Text(text.to_owned()), text.to_owned())
        },
        Scalar::Length(_) | Scalar::Color(_) | Scalar::List(_) => {
            return Err(refusal(
                Code::InvalidValue,
                "an ODS cell value is text, a number or a boolean",
            )
            .with_property("value"));
        },
    })
}

/// The stored OpenFormula token of an authored formula: `of:=` is added to
/// a bare expression or one with a leading `=`; a namespaced token
/// (`of:=…`, `oooc:=…`) is kept as written.
fn formula_token(text: &str) -> String {
    let namespaced = text.split_once(":=").is_some_and(|(prefix, _)| {
        !prefix.is_empty() && prefix.bytes().all(|b| b.is_ascii_alphanumeric())
    });
    if namespaced {
        text.to_owned()
    } else {
        format!("of:={}", text.strip_prefix('=').unwrap_or(text))
    }
}

/// Every logical coordinate of the cell and range targets, per sheet.
fn cell_coordinates(
    targets: &[Target],
    limits: &SemanticLimits,
) -> Result<BTreeMap<u32, Vec<(u32, u32)>>> {
    let mut sheets: BTreeMap<u32, Vec<(u32, u32)>> = BTreeMap::new();
    let mut total = 0_usize;
    for target in targets {
        match *target {
            Target::Cell { sheet, row, column } => {
                total = total.saturating_add(1);
                sheets.entry(sheet).or_default().push((row, column));
            },
            Target::Range {
                sheet,
                rect: (first_row, first_column, last_row, last_column),
            } => {
                let rows = (last_row - first_row) as usize + 1;
                let columns = (last_column - first_column) as usize + 1;
                let area = rows.saturating_mul(columns);
                total = total.saturating_add(area);
                if total > limits.max_expanded_targets() {
                    return Err(limit(
                        LimitKind::ExpandedTargets,
                        total,
                        limits.max_expanded_targets(),
                    ));
                }
                let coordinates = sheets.entry(sheet).or_default();
                coordinates.try_reserve(area).map_err(|_| {
                    refusal(
                        Code::InternalError,
                        "could not reserve the ODS range targets",
                    )
                })?;
                for row in first_row..=last_row {
                    for column in first_column..=last_column {
                        coordinates.push((row, column));
                    }
                }
            },
            _ => {
                return Err(refusal(
                    Code::UnsupportedElement,
                    "an ODS cell write target must be cells and ranges",
                ));
            },
        }
        if total > limits.max_expanded_targets() {
            return Err(limit(
                LimitKind::ExpandedTargets,
                total,
                limits.max_expanded_targets(),
            ));
        }
    }
    for coordinates in sheets.values_mut() {
        coordinates.sort_unstable();
        coordinates.dedup();
    }
    Ok(sheets)
}

/// The per-sheet cell replacements `assignments` make on `targets`.
fn cell_changes(
    file: &Ods,
    targets: &[Target],
    assignments: &CellAssignments,
    limits: &SemanticLimits,
) -> Result<BTreeMap<u32, BTreeMap<(u32, u32), Cell>>> {
    let sheets = file.worksheets()?;
    let mut changes: BTreeMap<u32, BTreeMap<(u32, u32), Cell>> = BTreeMap::new();
    let mut staged = 0_usize;
    for (sheet, coordinates) in cell_coordinates(targets, limits)? {
        let data = sheets
            .get(sheet as usize)
            .ok_or_else(super::node::missing_sheet)?;
        let entry = changes.entry(sheet).or_default();
        for (row, column) in coordinates {
            staged = staged.saturating_add(assignments.charge());
            if staged > limits.max_staged_bytes() {
                return Err(limit(
                    LimitKind::StagedBytes,
                    staged,
                    limits.max_staged_bytes(),
                ));
            }
            let existing = data.cell(row as usize, column as usize);
            if let Some(cell) = assignments.apply(existing)? {
                entry.insert((row, column), cell);
            }
        }
    }
    if let Some(merge) = assignments.merge {
        for target in targets {
            if let Target::Range { sheet, rect } = *target {
                let data = sheets
                    .get(sheet as usize)
                    .ok_or_else(super::node::missing_sheet)?;
                merge_range(data, rect, merge, changes.entry(sheet).or_default())?;
            }
        }
    }
    Ok(changes)
}

fn set_cells(
    file: &Ods,
    targets: &[Target],
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let ranges = targets
        .iter()
        .any(|target| matches!(target, Target::Range { .. }));
    let assignments = CellAssignments::admit(props, ranges)?;
    if let Some(style) = assignments.style.as_ref().and_then(Option::as_deref) {
        validate_style(file, style)?;
    }
    let changes = cell_changes(file, targets, &assignments, limits)?;
    if changes.values().all(BTreeMap::is_empty) {
        return no_op(file, limits);
    }
    let mut edit = file.leaf_edit()?;
    stage_cells(&mut edit, changes)?;
    finish(file, edit, 1, limits)
}

/// Stage per-sheet cell replacements through the worksheet family, in
/// batches the leaf admits.
pub(super) fn stage_cells(
    edit: &mut Edit,
    changes: BTreeMap<u32, BTreeMap<(u32, u32), Cell>>,
) -> Result<()> {
    edit.worksheets(|worksheets| {
        for (sheet, cells) in changes {
            let mut batch = Vec::new();
            for ((row, column), cell) in cells {
                batch.push(CellChange::new(row as usize, column as usize, cell));
                if batch.len() == litchi_ods::worksheet::MAX_CELL_CHANGES {
                    worksheets.set_cells(
                        Selector::Position(Position::new(sheet as usize)),
                        std::mem::take(&mut batch),
                    )?;
                }
            }
            if !batch.is_empty() {
                worksheets.set_cells(Selector::Position(Position::new(sheet as usize)), batch)?;
            }
        }
        Ok(())
    })
    .map_err(|error| odf_leaf(error, Code::InvalidValue))
}

/// Stage a merge (`true`) or unmerge (`false`) of `rect` into `cells`,
/// starting from the cells already staged there.
fn merge_range(
    sheet: &Sheet,
    rect: (u32, u32, u32, u32),
    merge: bool,
    cells: &mut BTreeMap<(u32, u32), Cell>,
) -> Result<()> {
    let (first_row, first_column, last_row, last_column) = rect;
    let current = |cells: &BTreeMap<(u32, u32), Cell>, at: (u32, u32)| {
        cells.get(&at).cloned().unwrap_or_else(|| {
            let mut cell = sheet
                .cell(at.0 as usize, at.1 as usize)
                .cloned()
                .unwrap_or_else(Cell::empty);
            cell.repeat = NonZeroUsize::MIN;
            cell
        })
    };
    let anchor = current(cells, (first_row, first_column));
    if !merge {
        let spans = matches!(anchor.merge, Merge::Span { rows, columns }
            if rows.get() as u64 == u64::from(last_row - first_row) + 1
                && columns.get() as u64 == u64::from(last_column - first_column) + 1);
        if !spans {
            return Ok(());
        }
    } else if (first_row, first_column) == (last_row, last_column) {
        return Err(refusal(
            Code::InvalidArgument,
            "a merged ODS range spans at least two cells",
        )
        .with_property("merge"));
    }
    for row in first_row..=last_row {
        for column in first_column..=last_column {
            let mut cell = current(cells, (row, column));
            if (row, column) == (first_row, first_column) {
                if merge {
                    cell.set_span(
                        (last_row - first_row) as usize + 1,
                        (last_column - first_column) as usize + 1,
                    )
                    .map_err(|error| leaf(error, Code::InvalidValue).with_property("merge"))?;
                } else {
                    cell.merge = Merge::None;
                }
            } else {
                // A covered cell carries no content (ODF 1.3 §9.1.5); merging
                // keeps the anchor's value, as spreadsheet applications do.
                if merge {
                    clear_content(&mut cell);
                }
                cell.set_covered(merge);
            }
            cells.insert((row, column), cell);
        }
    }
    Ok(())
}

/// Refuse a style name the package does not define.
fn validate_style(file: &Ods, name: &str) -> Result<()> {
    file.materialized()?
        .effective_cell_style(name)
        .map(|_| ())
        .map_err(|error| odf_leaf(error, Code::InvalidValue).with_property("style"))
}

fn single<T: Copy>(
    targets: &[Target],
    pick: impl Fn(Target) -> Option<T>,
    what: &'static str,
) -> Result<T> {
    match targets {
        [target] => pick(*target).ok_or_else(|| refusal(Code::UnsupportedElement, what)),
        _ => Err(refusal(Code::InvalidArgument, what)),
    }
}

fn set_sheet(
    file: &Ods,
    targets: &[Target],
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let position = single(
        targets,
        |target| match target {
            Target::Sheet(position) => Some(position),
            _ => None,
        },
        "an ODS sheet rename requires exactly one sheet target",
    )?;
    let mut name = None;
    for (key, value) in props.iter() {
        let (_, entry) = SHEET_WRITE_TABLE.resolve(key.as_str()).map_err(|error| {
            error
                .with_format(FileFormat::Ods)
                .with_property(key.as_str().to_owned())
        })?;
        let entry = *entry;
        entry
            .prepare(Operation::Set, value, |SheetSetter::Name, value| {
                let Some(Scalar::Text(value)) = value else {
                    return Err(
                        refusal(Code::InvalidValue, "an ODS sheet name requires text")
                            .with_property("name"),
                    );
                };
                name = Some(value.to_owned());
                Ok(())
            })
            .map_err(|error| error.with_format(FileFormat::Ods))?;
    }
    let Some(name) = name else {
        return no_op(file, limits);
    };
    let sheets = file.worksheets()?;
    let current = &sheets
        .get(position as usize)
        .ok_or_else(super::node::missing_sheet)?
        .name;
    if *current == name {
        return no_op(file, limits);
    }
    validate_sheet_name(file, Some(position as usize), &name)?;
    refuse_sheet_dependencies(file, current)?;
    // The leaf moves exact table fragments: copy the sheet under its new
    // name in its own slot, then drop the original after it.
    let mut edit = file.leaf_edit()?;
    edit.copy_sheet(current, &name, SheetPosition::Index(position as usize))
        .map_err(|error| odf_leaf(error, Code::UnsupportedCapability).with_property("name"))?;
    edit.remove_sheet(current)
        .map_err(|error| odf_leaf(error, Code::UnsupportedCapability).with_property("name"))?;
    finish(file, edit, 1, limits)
}

/// Refuse an invalid sheet name or one another sheet holds (ODF names
/// compare as LibreOffice does: case-insensitively).
fn validate_sheet_name(file: &Ods, except: Option<usize>, name: &str) -> Result<()> {
    let invalid = name.is_empty()
        || name.chars().count() > 255
        || name.starts_with('\'')
        || name.ends_with('\'')
        || name
            .chars()
            .any(|c| matches!(c, '[' | ']' | '*' | '?' | ':' | '/' | '\\') || c.is_control());
    if invalid {
        return Err(
            refusal(Code::InvalidValue, "the ODS sheet name is invalid").with_property("name")
        );
    }
    for (position, sheet) in file.worksheets()?.iter().enumerate() {
        if Some(position) != except
            && crate::path::vocabulary::Case::Caseless.matches(&sheet.name, name)
        {
            return Err(
                refusal(Code::DuplicateName, "a sheet name is already in use")
                    .with_property("name"),
            );
        }
    }
    Ok(())
}

/// Refuse renaming or removing a sheet that a formula, a named range or
/// expression still names: the leaf rewrites no reference.
fn refuse_sheet_dependencies(file: &Ods, name: &str) -> Result<()> {
    let named = |text: &str| text.contains(name);
    let in_formula = file.worksheets()?.iter().any(|sheet| {
        sheet.rows.iter().any(|row| {
            row.cells
                .iter()
                .any(|cell| cell.formula.as_deref().is_some_and(named))
        })
    });
    let in_definition = file.definitions()?.iter().any(|definition| {
        named(super::node::definition_reference(definition))
            || definition.scope().sheet_name() == Some(name)
    });
    if in_formula || in_definition {
        return Err(refusal(
            Code::UnsupportedCapability,
            "a formula or named range still references the sheet; the ODS writer rewrites no reference",
        )
        .with_suggestion("remove or retarget the references first"));
    }
    Ok(())
}

fn set_name(
    file: &Ods,
    targets: &[Target],
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let index = single(
        targets,
        |target| match target {
            Target::Name(index) => Some(index),
            _ => None,
        },
        "an ODS named-range set requires exactly one named range",
    )?;
    let definitions = file.definitions()?;
    let current = definitions.get(index).ok_or_else(|| {
        refusal(
            Code::PathNotFound,
            "the named range is outside the workbook",
        )
    })?;
    let mut updated = current.clone();
    for (key, value) in props.iter() {
        let (_, entry) = NAME_WRITE_TABLE.resolve(key.as_str()).map_err(|error| {
            error
                .with_format(FileFormat::Ods)
                .with_property(key.as_str().to_owned())
        })?;
        let entry = *entry;
        entry
            .prepare(Operation::Set, value, |setter, value| {
                let Some(Scalar::Text(text)) = value else {
                    return Err(refusal(
                        Code::InvalidValue,
                        "a named-range property requires text",
                    ));
                };
                match (setter, &mut updated) {
                    (NameSetter::Name, Definition::Range(range)) => range.name = text.to_owned(),
                    (NameSetter::Name, Definition::Expression(expression)) => {
                        expression.name = text.to_owned();
                    },
                    (NameSetter::Ref, Definition::Range(range)) => {
                        range.cell_range_address = range_address(text);
                    },
                    (NameSetter::Ref, Definition::Expression(expression)) => {
                        expression.expression = text.to_owned();
                    },
                    (NameSetter::Scope, _) => {
                        return Err(refusal(
                            Code::UnsupportedCapability,
                            "a named range's scope is fixed once added",
                        ));
                    },
                }
                Ok(())
            })
            .map_err(|error| error.with_format(FileFormat::Ods))?;
    }
    if updated == *current {
        return no_op(file, limits);
    }
    let mut edit = file.leaf_edit()?;
    edit.definitions(|definitions| {
        definitions
            .replace(Position::new(index), updated)
            .map(|_| ())
    })
    .map_err(|error| odf_leaf(error, Code::InvalidValue))?;
    finish(file, edit, 1, limits)
}

/// An ODF cell range address from an Excel-style one (`Sheet1!$A$1:$B$2`
/// becomes `$Sheet1.$A$1:.$B$2`); an ODF address is kept as written.
fn range_address(text: &str) -> String {
    let Some((sheet, cells)) = text.split_once('!') else {
        return text.to_owned();
    };
    let sheet = sheet.trim_matches('\'');
    let mut out = format!("${sheet}.");
    match cells.split_once(':') {
        Some((first, last)) => {
            out.push_str(first);
            out.push_str(":.");
            out.push_str(last);
        },
        None => out.push_str(cells),
    }
    out
}

fn set_comment(
    file: &Ods,
    targets: &[Target],
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let index = single(
        targets,
        |target| match target {
            Target::Comment(index) => Some(index),
            _ => None,
        },
        "an ODS annotation set requires exactly one annotation",
    )?;
    let annotations = super::node::annotations(file)?;
    let current = annotations
        .entries()
        .get(index)
        .ok_or_else(|| refusal(Code::PathNotFound, "the annotation is outside the workbook"))?
        .annotation();
    let mut updated = current.clone();
    apply_comment(&mut updated, props, Operation::Set, &mut None)?;
    if updated == *current {
        return no_op(file, limits);
    }
    let mut edit = file.leaf_edit()?;
    edit.annotations(|transaction| transaction.replace(index, updated))
        .map_err(|error| odf_leaf(error, Code::InvalidValue))?;
    finish(file, edit, 1, limits)
}

/// Apply annotation properties; `ref` (add only) lands in `reference`.
fn apply_comment(
    annotation: &mut Annotation,
    props: &Props<'_>,
    operation: Operation,
    reference: &mut Option<String>,
) -> Result<()> {
    for (key, value) in props.iter() {
        let (_, entry) = COMMENT_WRITE_TABLE.resolve(key.as_str()).map_err(|error| {
            error
                .with_format(FileFormat::Ods)
                .with_property(key.as_str().to_owned())
        })?;
        let entry = *entry;
        entry
            .prepare(operation, value, |setter, value| {
                let text = match value {
                    Some(Scalar::Text(text)) => Some(text),
                    None => None,
                    Some(_) => {
                        return Err(refusal(
                            Code::InvalidValue,
                            "an annotation property requires text",
                        ));
                    },
                };
                match setter {
                    CommentSetter::Ref => *reference = text.map(str::to_owned),
                    CommentSetter::Author => annotation.set_creator(text),
                    CommentSetter::Date => annotation.set_date(text),
                    CommentSetter::Text => {
                        annotation.clear_body();
                        for line in text.unwrap_or_default().split('\n') {
                            annotation.push_paragraph(line);
                        }
                    },
                }
                Ok(())
            })
            .map_err(|error| error.with_format(FileFormat::Ods))?;
    }
    annotation
        .validate()
        .map_err(|error| leaf(error, Code::InvalidValue))
}

// ---------------------------------------------------------------- remove

fn remove_targets(file: &Ods, targets: &[Target], limits: &SemanticLimits) -> Result<Commit> {
    let Some(first) = targets.first() else {
        return no_op(file, limits);
    };
    match first.element() {
        Kind::Cell | Kind::Range => remove_cells(file, targets, limits),
        Kind::Sheet => {
            let position = single(
                targets,
                |target| match target {
                    Target::Sheet(position) => Some(position),
                    _ => None,
                },
                "an ODS sheet removal takes one sheet per call",
            )?;
            let sheets = file.worksheets()?;
            if sheets.len() <= 1 {
                return Err(refusal(
                    Code::UnsupportedCapability,
                    "a spreadsheet keeps at least one sheet",
                ));
            }
            let name = &sheets
                .get(position as usize)
                .ok_or_else(super::node::missing_sheet)?
                .name;
            refuse_sheet_dependencies(file, name)?;
            let mut edit = file.leaf_edit()?;
            edit.remove_sheet(name)
                .map_err(|error| odf_leaf(error, Code::UnsupportedCapability))?;
            finish(file, edit, 1, limits)
        },
        Kind::Row => {
            let (sheet, index) = single(
                targets,
                |target| match target {
                    Target::Row { sheet, index } => Some((sheet, index)),
                    _ => None,
                },
                "an ODS row removal takes one row per call",
            )?;
            let name = sheet_name(file, sheet)?;
            let mut edit = file.leaf_edit()?;
            edit.edit_logical_rows(
                &name,
                &[LogicalRowOperation::Remove {
                    at: index as usize,
                    count: 1,
                }],
            )
            .map_err(|error| odf_leaf(error, Code::UnsupportedCapability))?;
            finish(file, edit, 1, limits)
        },
        Kind::NamedRange => {
            let index = single(
                targets,
                |target| match target {
                    Target::Name(index) => Some(index),
                    _ => None,
                },
                "an ODS named-range removal takes one name per call",
            )?;
            let mut edit = file.leaf_edit()?;
            edit.definitions(|definitions| definitions.remove(Position::new(index)).map(|_| ()))
                .map_err(|error| odf_leaf(error, Code::UnsupportedCapability))?;
            finish(file, edit, 1, limits)
        },
        Kind::Comment => {
            let mut indexes: Vec<usize> = targets
                .iter()
                .filter_map(|target| match target {
                    Target::Comment(index) => Some(*index),
                    _ => None,
                })
                .collect();
            if indexes.len() != targets.len() {
                return Err(refusal(
                    Code::InvalidArgument,
                    "an annotation removal query must select annotations only",
                ));
            }
            indexes.sort_unstable();
            indexes.dedup();
            let mut edit = file.leaf_edit()?;
            edit.annotations(|transaction| {
                // Remove from the end so earlier indexes stay valid.
                for index in indexes.iter().rev() {
                    transaction.remove(*index)?;
                }
                Ok(())
            })
            .map_err(|error| odf_leaf(error, Code::UnsupportedCapability))?;
            finish(file, edit, indexes.len(), limits)
        },
        Kind::MergeCell => {
            let sheets = file.worksheets()?;
            let mut changes: BTreeMap<u32, BTreeMap<(u32, u32), Cell>> = BTreeMap::new();
            for target in targets {
                let Target::Merge { sheet, rect } = *target else {
                    return Err(refusal(
                        Code::InvalidArgument,
                        "a merge removal query must select merge spans only",
                    ));
                };
                let data = sheets
                    .get(sheet as usize)
                    .ok_or_else(super::node::missing_sheet)?;
                merge_range(data, rect, false, changes.entry(sheet).or_default())?;
            }
            let mut edit = file.leaf_edit()?;
            stage_cells(&mut edit, changes)?;
            finish(file, edit, 1, limits)
        },
        Kind::Column => Err(refusal(
            Code::UnsupportedCapability,
            "removing an ODS column needs a logical column editor the leaf does not have",
        )
        .with_suggestion("follow-up: ODS adapter (record 0277)")),
        _ => Err(refusal(
            Code::UnsupportedElement,
            "ODS removes cells, ranges, sheets, rows, named ranges, annotations and merge spans",
        )),
    }
}

fn remove_cells(file: &Ods, targets: &[Target], limits: &SemanticLimits) -> Result<Commit> {
    let sheets = file.worksheets()?;
    let mut changes: BTreeMap<u32, BTreeMap<(u32, u32), Cell>> = BTreeMap::new();
    for (sheet, coordinates) in cell_coordinates(targets, limits)? {
        let data = sheets
            .get(sheet as usize)
            .ok_or_else(super::node::missing_sheet)?;
        for (row, column) in coordinates {
            let Some(existing) = data.cell(row as usize, column as usize) else {
                continue;
            };
            let mut cell = existing.clone();
            cell.repeat = NonZeroUsize::MIN;
            clear_content(&mut cell);
            cell.style_name = None;
            let mut same = existing.clone();
            same.repeat = NonZeroUsize::MIN;
            if same != cell {
                changes
                    .entry(sheet)
                    .or_default()
                    .insert((row, column), cell);
            }
        }
    }
    if changes.is_empty() {
        return no_op(file, limits);
    }
    let mut edit = file.leaf_edit()?;
    stage_cells(&mut edit, changes)?;
    finish(file, edit, 1, limits)
}

fn sheet_name(file: &Ods, position: u32) -> Result<String> {
    Ok(file
        .worksheets()?
        .get(position as usize)
        .ok_or_else(super::node::missing_sheet)?
        .name
        .clone())
}

// ---------------------------------------------------------------- add

fn sheet_path(position: usize) -> Path<'static> {
    Path::from_steps([Step::new("sheet").with(Predicate::Index(Position::new(position)))])
}

fn add(
    file: &Ods,
    parent: Target,
    node: &NewNode<'_>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    let kind = node.kind();
    if let Some((feature, implied)) = super::feature_write::feature_kind(kind, node.props()) {
        let parent = match parent {
            Target::Sheet(sheet) => super::feature_write::Parent {
                sheet,
                reference: None,
            },
            Target::Cell { sheet, row, column } => super::feature_write::Parent {
                sheet,
                reference: Some((row, column, row, column)),
            },
            Target::Range { sheet, rect } => super::feature_write::Parent {
                sheet,
                reference: Some(rect),
            },
            _ => {
                return Err(refusal(
                    Code::UnsupportedType,
                    "worksheet features are added under a worksheet, a cell or a range",
                ));
            },
        };
        return super::feature_write::add(file, &parent, feature, implied, node, limits);
    }
    if kind.eq_ignore_ascii_case("sheet") || kind.eq_ignore_ascii_case("worksheet") {
        return add_sheet(file, parent, node, limits);
    }
    if kind.eq_ignore_ascii_case("row") {
        return add_row(file, parent, node, limits);
    }
    if kind.eq_ignore_ascii_case("namedrange") || kind.eq_ignore_ascii_case("definedname") {
        return add_name(file, parent, node, limits);
    }
    if kind.eq_ignore_ascii_case("comment") || kind.eq_ignore_ascii_case("note") {
        return add_comment(file, parent, node, limits);
    }
    Err(refusal(
        Code::UnsupportedElement,
        "ODS adds sheets, rows, named ranges and annotations; set cells instead of adding them",
    ))
}

fn add_sheet(
    file: &Ods,
    parent: Target,
    node: &NewNode<'_>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    if parent != Target::Workbook {
        return Err(refusal(
            Code::UnsupportedType,
            "ODS sheets live directly under the workbook",
        ));
    }
    let mut name = None;
    for (key, value) in node.props().iter() {
        let (_, entry) = SHEET_WRITE_TABLE.resolve(key.as_str()).map_err(|error| {
            error
                .with_format(FileFormat::Ods)
                .with_property(key.as_str().to_owned())
        })?;
        let entry = *entry;
        entry
            .prepare(Operation::Add, value, |SheetSetter::Name, value| {
                let Some(Scalar::Text(value)) = value else {
                    return Err(
                        refusal(Code::InvalidValue, "an ODS sheet name requires text")
                            .with_property("name"),
                    );
                };
                name = Some(value.to_owned());
                Ok(())
            })
            .map_err(|error| error.with_format(FileFormat::Ods))?;
    }
    let count = file.worksheet_count()?;
    let name = match name {
        Some(name) => name,
        None => {
            let mut number = count + 1;
            loop {
                let candidate = format!("Sheet{number}");
                if validate_sheet_name(file, None, &candidate).is_ok() {
                    break candidate;
                }
                number += 1;
            }
        },
    };
    validate_sheet_name(file, None, &name)?;
    let position = match node.placement() {
        None | Some(Placement::End) => count,
        Some(Placement::Start) => 0,
        Some(Placement::Position(position)) => position.get(),
        Some(Placement::Before(path)) => sheet_target(file, path)?,
        Some(Placement::After(path)) => sheet_target(file, path)?.saturating_add(1),
        Some(Placement::Into(_, position)) => position.map_or(count, Position::get),
    };
    if position > count {
        return Err(refusal(
            Code::OutOfRange,
            "the ODS sheet insertion position is outside the workbook",
        ));
    }
    if limits.max_operations() < 1 {
        return Err(limit(LimitKind::Operations, 1, limits.max_operations()));
    }
    // One empty row with one empty cell, as spreadsheet applications author
    // a new sheet: later cell writes splice rows next to it.
    let mut sheet = Sheet::new(name.clone()).map_err(|error| leaf(error, Code::InvalidValue))?;
    let mut row = Row::new();
    row.push_cell(Cell::empty())
        .map_err(|error| leaf(error, Code::InvalidValue))?;
    sheet
        .push_row(row)
        .map_err(|error| leaf(error, Code::InvalidValue))?;
    let mut edit = file.leaf_edit()?;
    edit.append_sheet(&sheet)
        .map_err(|error| odf_leaf(error, Code::UnsupportedCapability))?;
    if position != count {
        edit.move_sheet(&name, SheetPosition::Index(position))
            .map_err(|error| odf_leaf(error, Code::UnsupportedCapability))?;
    }
    Ok((finish(file, edit, 1, limits)?, sheet_path(position)))
}

/// The position of the sheet a placement path names.
fn sheet_target(file: &Ods, path: &PathArg<'_>) -> Result<usize> {
    match Target::of(&crate::path::resolve::get(&Nav::new(file), path)?) {
        Target::Sheet(position) => Ok(position as usize),
        _ => Err(refusal(
            Code::UnsupportedType,
            "an ODS sheet placement names a sheet",
        )),
    }
}

/// The zero-based row a placement path names in `sheet`.
fn row_target(file: &Ods, sheet: u32, path: &PathArg<'_>) -> Result<u32> {
    match Target::of(&crate::path::resolve::get(&Nav::new(file), path)?) {
        Target::Row {
            sheet: owner,
            index,
        } if owner == sheet => Ok(index),
        _ => Err(refusal(
            Code::UnsupportedType,
            "an ODS row placement names a row of the same sheet",
        )),
    }
}

/// The row after the last content row: where `End` inserts.
fn row_end(file: &Ods, sheet: u32) -> Result<u32> {
    Ok(super::node::sheet_index(file, sheet)?
        .row_entries()
        .last()
        .map_or(0, |(row, _)| row + 1))
}

fn add_row(
    file: &Ods,
    parent: Target,
    node: &NewNode<'_>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    let Target::Sheet(sheet) = parent else {
        return Err(refusal(
            Code::UnsupportedType,
            "ODS rows are added under a sheet",
        ));
    };
    if !node.props().is_empty() {
        return Err(refusal(
            Code::UnsupportedProperty,
            "an added ODS row takes no properties; set its cells afterwards",
        ));
    }
    let at = match node.placement() {
        None | Some(Placement::End) => row_end(file, sheet)?,
        Some(Placement::Start) => 0,
        Some(Placement::Position(position)) => u32::try_from(position.get()).unwrap_or(u32::MAX),
        Some(Placement::Before(path)) => row_target(file, sheet, path)?,
        Some(Placement::After(path)) => row_target(file, sheet, path)?.saturating_add(1),
        Some(Placement::Into(_, position)) => match position {
            Some(position) => u32::try_from(position.get()).unwrap_or(u32::MAX),
            None => row_end(file, sheet)?,
        },
    };
    let name = sheet_name(file, sheet)?;
    let mut edit = file.leaf_edit()?;
    edit.edit_logical_rows(
        &name,
        &[LogicalRowOperation::Insert {
            at: at as usize,
            rows: vec![Row::new()],
        }],
    )
    .map_err(|error| odf_leaf(error, Code::UnsupportedCapability))?;
    let path = Path::from_steps([
        Step::new("sheet").with(Predicate::Index(Position::new(sheet as usize))),
        Step::new("row").with(Predicate::Index(Position::new(at as usize))),
    ]);
    Ok((finish(file, edit, 1, limits)?, path))
}

fn add_name(
    file: &Ods,
    parent: Target,
    node: &NewNode<'_>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    if parent != Target::Workbook {
        return Err(refusal(
            Code::UnsupportedType,
            "ODS named ranges live under the workbook",
        ));
    }
    let (mut name, mut reference, mut scope) = (None, None, Scope::Global);
    for (key, value) in node.props().iter() {
        let (_, entry) = NAME_WRITE_TABLE.resolve(key.as_str()).map_err(|error| {
            error
                .with_format(FileFormat::Ods)
                .with_property(key.as_str().to_owned())
        })?;
        let entry = *entry;
        entry
            .prepare(Operation::Add, value, |setter, value| {
                let Some(Scalar::Text(text)) = value else {
                    return Err(refusal(
                        Code::InvalidValue,
                        "a named-range property requires text",
                    ));
                };
                match setter {
                    NameSetter::Name => name = Some(text.to_owned()),
                    NameSetter::Ref => reference = Some(range_address(text)),
                    NameSetter::Scope => {
                        if !text.eq_ignore_ascii_case("workbook") {
                            scope = Scope::sheet(text);
                        }
                    },
                }
                Ok(())
            })
            .map_err(|error| error.with_format(FileFormat::Ods))?;
    }
    let (Some(name), Some(reference)) = (name, reference) else {
        return Err(refusal(
            Code::MissingProperty,
            "an ODS named range needs `name` and `ref`",
        ));
    };
    let definition = Definition::Range(
        NamedRange::new(name, reference, scope).map_err(|error| leaf(error, Code::InvalidValue))?,
    );
    let index = file.definitions()?.len();
    let mut edit = file.leaf_edit()?;
    edit.definitions(|definitions| definitions.add(definition))
        .map_err(|error| odf_leaf(error, Code::InvalidValue))?;
    let path =
        Path::from_steps([Step::new("namedrange").with(Predicate::Index(Position::new(index)))]);
    Ok((finish(file, edit, 1, limits)?, path))
}

fn add_comment(
    file: &Ods,
    parent: Target,
    node: &NewNode<'_>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    let mut annotation = Annotation::new("");
    annotation.clear_body();
    let mut reference = None;
    apply_comment(
        &mut annotation,
        node.props(),
        Operation::Add,
        &mut reference,
    )?;
    let (sheet, row, column) = match (parent, reference) {
        (Target::Cell { sheet, row, column }, None) => (sheet, row, column),
        (Target::Sheet(sheet), Some(reference)) => {
            let cell = litchi_core::selector::CellRef::parse(&reference).ok_or_else(|| {
                refusal(Code::InvalidValue, "the annotation `ref` is not an A1 cell")
                    .with_property("ref")
            })?;
            (sheet, cell.row(), cell.column())
        },
        (Target::Cell { .. }, Some(_)) => {
            return Err(refusal(
                Code::InvalidArgument,
                "an annotation added under a cell takes no `ref`",
            ));
        },
        _ => {
            return Err(refusal(
                Code::UnsupportedType,
                "ODS annotations are added under a cell, or under a sheet with `ref`",
            ));
        },
    };
    let name = sheet_name(file, sheet)?;
    let mut edit = file.leaf_edit()?;
    edit.annotations(|transaction| {
        transaction.add(&name, row as usize, column as usize, annotation)
    })
    .map_err(|error| odf_leaf(error, Code::UnsupportedCapability))?;
    let commit = finish(file, edit, 1, limits)?;
    // The new annotation's ordinal among its sheet's annotations, in
    // document order of the committed snapshot.
    let ordinal = super::node::annotations(&commit.after)?
        .entries()
        .iter()
        .filter(|entry| entry.cell().sheet() == name)
        .position(|entry| {
            entry.cell().row() == row as usize && entry.cell().column() == column as usize
        })
        .unwrap_or_default();
    let path = Path::from_steps([
        Step::new("sheet").with(Predicate::Index(Position::new(sheet as usize))),
        Step::new("comment").with(Predicate::Index(Position::new(ordinal))),
    ]);
    Ok((commit, path))
}

// ---------------------------------------------------------------- move / swap

fn move_target(
    file: &Ods,
    target: Target,
    placement: &Placement<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    match target {
        Target::Sheet(position) => {
            let count = file.worksheet_count()?;
            let from = position as usize;
            // Positional destinations are measured after detaching the sheet.
            let to = match placement {
                Placement::End => count - 1,
                Placement::Start => 0,
                Placement::Position(position) => position.get(),
                Placement::Before(path) => {
                    let before = sheet_target(file, path)?;
                    if before > from { before - 1 } else { before }
                },
                Placement::After(path) => {
                    let after = sheet_target(file, path)?;
                    if after >= from { after } else { after + 1 }
                },
                Placement::Into(_, position) => position.map_or(count - 1, Position::get),
            };
            if to >= count {
                return Err(refusal(
                    Code::OutOfRange,
                    "the ODS sheet destination is outside the workbook",
                ));
            }
            if to == from {
                return no_op(file, limits);
            }
            let name = sheet_name(file, position)?;
            let mut edit = file.leaf_edit()?;
            edit.move_sheet(&name, SheetPosition::Index(to))
                .map_err(|error| odf_leaf(error, Code::UnsupportedCapability))?;
            finish(file, edit, 1, limits)
        },
        Target::Row { sheet, index } => {
            let end = row_end(file, sheet)?;
            let to = match placement {
                // The end of the content rows, measured after detaching.
                Placement::End if index < end => end - 1,
                Placement::End => end,
                Placement::Start => 0,
                Placement::Position(position) => u32::try_from(position.get()).unwrap_or(u32::MAX),
                Placement::Before(path) => {
                    let before = row_target(file, sheet, path)?;
                    if before > index { before - 1 } else { before }
                },
                Placement::After(path) => {
                    let after = row_target(file, sheet, path)?;
                    if after >= index { after } else { after + 1 }
                },
                Placement::Into(_, _) => {
                    return Err(refusal(
                        Code::UnsupportedType,
                        "ODS rows move within their sheet",
                    ));
                },
            };
            if to == index {
                return no_op(file, limits);
            }
            let name = sheet_name(file, sheet)?;
            let mut edit = file.leaf_edit()?;
            edit.edit_logical_rows(
                &name,
                &[LogicalRowOperation::Move {
                    at: index as usize,
                    count: 1,
                    to: to as usize,
                }],
            )
            .map_err(|error| odf_leaf(error, Code::UnsupportedCapability))?;
            finish(file, edit, 1, limits)
        },
        _ => Err(refusal(
            Code::UnsupportedElement,
            "ODS moves sheets and rows",
        )),
    }
}

fn swap_targets(
    file: &Ods,
    left: Target,
    right: Target,
    limits: &SemanticLimits,
) -> Result<Commit> {
    match (left, right) {
        (Target::Sheet(left), Target::Sheet(right)) => {
            if left == right {
                return no_op(file, limits);
            }
            let (first, second) = (left.min(right) as usize, left.max(right) as usize);
            let first_name = sheet_name(file, first as u32)?;
            let second_name = sheet_name(file, second as u32)?;
            let mut edit = file.leaf_edit()?;
            edit.move_sheet(&second_name, SheetPosition::Index(first))
                .map_err(|error| odf_leaf(error, Code::UnsupportedCapability))?;
            edit.move_sheet(&first_name, SheetPosition::Index(second))
                .map_err(|error| odf_leaf(error, Code::UnsupportedCapability))?;
            finish(file, edit, 1, limits)
        },
        (
            Target::Row { sheet, index: left },
            Target::Row {
                sheet: other,
                index: right,
            },
        ) if sheet == other => {
            if left == right {
                return no_op(file, limits);
            }
            let (first, second) = (left.min(right) as usize, left.max(right) as usize);
            let name = sheet_name(file, sheet)?;
            let mut edit = file.leaf_edit()?;
            edit.edit_logical_rows(
                &name,
                &[
                    LogicalRowOperation::Move {
                        at: second,
                        count: 1,
                        to: first,
                    },
                    LogicalRowOperation::Move {
                        at: first + 1,
                        count: 1,
                        to: second,
                    },
                ],
            )
            .map_err(|error| odf_leaf(error, Code::UnsupportedCapability))?;
            finish(file, edit, 1, limits)
        },
        _ => Err(refusal(
            Code::UnsupportedElement,
            "ODS swaps two sheets, or two rows of one sheet",
        )),
    }
}

#[cfg(test)]
mod tests {
    use super::{formula_token, leaf, odf_leaf, range_address};
    use crate::error::{Code, Source};

    #[test]
    fn typed_odf_errors_keep_their_classification() {
        let refusal = odf_leaf(
            litchi_odf_common::Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedTarget,
                "set_feature",
                "the target does not accept this feature",
            ),
            Code::InvalidValue,
        );
        assert_eq!(refusal.code(), Code::UnsupportedType);
        assert!(matches!(refusal.leaf(), Source::Odf(_)));

        let conflict = odf_leaf(
            litchi_odf_common::Error::patch_conflict("set_feature", "the package base differs"),
            Code::UnsupportedCapability,
        );
        assert_eq!(conflict.code(), Code::InvalidInput);
        assert!(matches!(conflict.leaf(), Source::Odf(_)));
    }

    #[test]
    fn typed_core_refusals_keep_their_classification() {
        let refusal = leaf(
            litchi_core::Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "set_feature",
                "the operation is unavailable",
            )),
            Code::InvalidValue,
        );
        assert_eq!(refusal.code(), Code::UnsupportedCapability);
        assert!(matches!(refusal.leaf(), Source::Core(_)));
    }

    #[test]
    fn formulas_gain_the_openformula_namespace() {
        assert_eq!(formula_token("SUM([.A1:.A2])"), "of:=SUM([.A1:.A2])");
        assert_eq!(formula_token("=[.A1]*2"), "of:=[.A1]*2");
        assert_eq!(formula_token("of:=[.A1]"), "of:=[.A1]");
        assert_eq!(formula_token("oooc:=[.A1]"), "oooc:=[.A1]");
    }

    #[test]
    fn excel_range_addresses_become_odf_addresses() {
        assert_eq!(range_address("Sheet1!$A$1:$B$2"), "$Sheet1.$A$1:.$B$2");
        assert_eq!(range_address("'My Sheet'!$C$3"), "$My Sheet.$C$3");
        assert_eq!(range_address("$Sheet1.$A$1"), "$Sheet1.$A$1");
    }
}
