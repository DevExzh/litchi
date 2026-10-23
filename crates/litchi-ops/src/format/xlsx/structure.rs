//! Structural XLSX sheet verbs (record 0102).
//!
//! The unified structural surface is intentionally limited to workbook sheet
//! tabs.  Sheet names, tab order and removal dependencies remain owned by the
//! native workbook transaction; this adapter only resolves the public path and
//! property contracts and hands the native commit to the ordinary XLSX write
//! result.  No XML is rebuilt by the unified layer.

use litchi_core::selector::{Path, Predicate, Step};
use litchi_core::{FileFormat, Position};
use litchi_xlsx as xlsx;

use super::node::Node;
use super::spec::Kind;
use super::write::{Commit, edit_with_snapshot_limit};
use super::{Nav, Xlsx};
use crate::error::{Code, Error, Result};
use crate::new_node::NewNode;
use crate::path::PathArg;
use crate::placement::Placement;
use crate::registry::Operation;
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};

/// Add a worksheet, remove one, or change the order of workbook sheet tabs.
///
/// The native transaction is the source of truth for the final catalog.  In
/// particular, relationship checks and active/visible-tab invariants are not
/// reproduced here.
impl Xlsx {
    /// Add a worksheet below the workbook root and return its canonical path.
    pub(crate) fn add<'p, 'n>(
        &self,
        parent: PathArg<'p>,
        node: NewNode<'n>,
    ) -> Result<(Commit, Path<'static>)> {
        self.add_with_limits(parent, node, &SemanticLimits::default())
    }

    pub(crate) fn add_with_limits<'p, 'n>(
        &self,
        parent: PathArg<'p>,
        node: NewNode<'n>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        self.0
            .owner
            .finish(|| self.add_inner(parent, node, limits))
            .and_then(|(commit, path)| {
                self.enforce_signed_noop(commit)
                    .map(|commit| (commit, path))
            })
    }

    fn add_inner<'p, 'n>(
        &self,
        parent: PathArg<'p>,
        node: NewNode<'n>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        // Admit the owning editor first.  This keeps retention, signatures and
        // encryption policy ahead of authored property conversion, matching
        // the cell and DOCX write adapters.
        let source = self.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        // The element kind first (record 0269, finding X1): a kind the
        // workbook cannot add is `unsupported_element` whatever its parent;
        // an addable kind under the wrong parent is `unsupported_type`.
        let kind = new_kind(node.kind(), node.props())?;
        let parent = resolve_path(self, &parent)?;
        if matches!(kind, NewKind::Chart) {
            drop(source);
            return self.add_chart(&parent, node, limits);
        }
        if let NewKind::Feature(kind, implied) = kind {
            drop(source);
            return self.add_feature(&parent, kind, implied, node, limits);
        }
        // A worksheet cell and a rich-text run (record 0338).
        if matches!(kind, NewKind::Cell) {
            drop(source);
            return self.add_cell(&parent, &node, limits);
        }
        if matches!(kind, NewKind::Run) {
            drop(source);
            return self.add_run_node(&parent, &node, limits);
        }
        if parent.element() != Kind::Workbook {
            return Err(Error::new(
                Code::UnsupportedType,
                "an XLSX sheet can only be added directly under the workbook",
            )
            .with_format(FileFormat::Xlsx));
        }
        let placement = node.placement().cloned();
        let properties = super::sheet_write::admit_sheet_properties(Operation::Add, node.props())?;
        let name = match properties.name {
            Some(name) => name,
            None => {
                let next = self.0.owner.len().checked_add(1).ok_or_else(|| {
                    Error::new(
                        Code::OutOfRange,
                        "the XLSX sheet catalog position is outside its supported range",
                    )
                    .with_format(FileFormat::Xlsx)
                })?;
                let authored = format!("Sheet{next}");
                xlsx::sheet::Name::new(&authored).map_err(|_| {
                    Error::new(
                        Code::InternalError,
                        "the XLSX default sheet name could not be validated",
                    )
                    .with_format(FileFormat::Xlsx)
                })?
            },
        };
        let mut edit = edit_with_snapshot_limit(&source, limits.max_snapshot_bytes())?;
        let position = add_at(
            &mut edit,
            name,
            placement.as_ref(),
            properties.hidden.unwrap_or(false),
            self,
        )?;
        let native = edit.commit().map_err(native_error)?;
        let commit =
            Commit::from_native_with_limit(self, source, native, limits.max_snapshot_bytes())?;
        Ok((commit, sheet_path(position)))
    }

    /// Remove one worksheet selected by a structural path.
    pub(crate) fn remove<'p>(&self, target: TargetArg<'p>) -> Result<Commit> {
        self.remove_with_limits(target, &SemanticLimits::default())
    }

    /// Remove one or more sheet tabs selected by a bounded structural path or
    /// scoped query in one native workbook transaction.
    pub(crate) fn remove_with_limits<'p>(
        &self,
        target: TargetArg<'p>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        self.0
            .owner
            .finish(|| self.remove_inner(target, limits))
            .and_then(|commit| self.enforce_signed_noop(commit))
    }

    fn remove_inner<'p>(&self, target: TargetArg<'p>, limits: &SemanticLimits) -> Result<Commit> {
        if let Some(commit) = self.try_remove_run(&target, limits)? {
            return Ok(commit);
        }
        if let Some(commit) = self.try_remove_chart(&target, limits)? {
            return Ok(commit);
        }
        if let Some(commit) = self.try_remove_feature(&target, limits)? {
            return Ok(commit);
        }
        let source = self.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        let positions = super::sheet_write::resolve_sheet_targets(self, &target, limits)?
            .ok_or_else(unsupported_removal_target)?;
        check_removal_limits(positions.len(), limits)?;
        self.check_source()?;
        let mut edit = edit_with_snapshot_limit(&source, limits.max_snapshot_bytes())?;
        for position in positions {
            if edit.remove(position).map_err(native_error)?.is_none() {
                return Err(Error::new(
                    Code::NotFound,
                    "the XLSX sheet was not available for removal",
                )
                .with_format(FileFormat::Xlsx));
            }
        }
        let native = edit.commit().map_err(native_error)?;
        Commit::from_native_with_limit(self, source, native, limits.max_snapshot_bytes())
    }

    /// Move one sheet tab to a checked placement.
    pub(crate) fn move_to<'p, 'q>(
        &self,
        path: PathArg<'p>,
        placement: Placement<'q>,
    ) -> Result<Commit> {
        self.move_to_with_limits(path, placement, &SemanticLimits::default())
    }

    pub(crate) fn move_to_with_limits<'p, 'q>(
        &self,
        path: PathArg<'p>,
        placement: Placement<'q>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        self.0
            .owner
            .finish(|| self.move_inner(path, placement, limits))
            .and_then(|commit| self.enforce_signed_noop(commit))
    }

    fn move_inner<'p, 'q>(
        &self,
        path: PathArg<'p>,
        placement: Placement<'q>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        let source = self.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        let node = resolve_path(self, &path)?;
        let source_position = sheet_position(&node)?;
        let destination = move_destination(self, &placement)?;
        self.check_source()?;
        let mut edit = edit_with_snapshot_limit(&source, limits.max_snapshot_bytes())?;
        let moved = match destination {
            MoveDestination::Position(position) => edit
                .move_to(source_position, position)
                .map_err(native_error)?,
            MoveDestination::Before(anchor) => edit
                .move_before(source_position, anchor)
                .map_err(native_error)?,
            MoveDestination::After(anchor) => edit
                .move_after(source_position, anchor)
                .map_err(native_error)?,
        };
        if moved.is_none() {
            return Err(Error::new(
                Code::NotFound,
                "the XLSX sheet was not available for reordering",
            )
            .with_format(FileFormat::Xlsx));
        }
        let native = edit.commit().map_err(native_error)?;
        Commit::from_native_with_limit(self, source, native, limits.max_snapshot_bytes())
    }

    /// Exchange two sheet tabs while retaining one native transaction.
    pub(crate) fn swap<'p, 'q>(&self, left: PathArg<'p>, right: PathArg<'q>) -> Result<Commit> {
        self.swap_with_limits(left, right, &SemanticLimits::default())
    }

    pub(crate) fn swap_with_limits<'p, 'q>(
        &self,
        left: PathArg<'p>,
        right: PathArg<'q>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        self.0
            .owner
            .finish(|| self.swap_inner(left, right, limits))
            .and_then(|commit| self.enforce_signed_noop(commit))
    }

    fn swap_inner<'p, 'q>(
        &self,
        left: PathArg<'p>,
        right: PathArg<'q>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        let source = self.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        let left = sheet_position(&resolve_path(self, &left)?)?;
        let right = sheet_position(&resolve_path(self, &right)?)?;
        self.check_source()?;
        let mut edit = edit_with_snapshot_limit(&source, limits.max_snapshot_bytes())?;
        if left != right {
            // Each `move_to` resolves its source identity against the
            // immutable base snapshot and applies its destination to the
            // projected order. Moving left to right first shifts the right
            // identity by one when necessary; the second move then places it
            // at the left slot, giving a true exchange even for non-adjacent
            // tabs.
            edit.move_to(left, right)
                .map_err(native_error)?
                .ok_or_else(|| not_found("the first XLSX sheet disappeared"))?;
            edit.move_to(right, left)
                .map_err(native_error)?
                .ok_or_else(|| not_found("the second XLSX sheet disappeared"))?;
        }
        let native = edit.commit().map_err(native_error)?;
        Commit::from_native_with_limit(self, source, native, limits.max_snapshot_bytes())
    }

    fn enforce_signed_noop(&self, commit: Commit) -> Result<Commit> {
        if self.signature() == super::Signature::Signed && !commit.is_noop_for(self) {
            Err(Error::signed(FileFormat::Xlsx))
        } else {
            Ok(commit)
        }
    }
}

#[derive(Clone, Copy)]
enum NewKind {
    Sheet,
    /// A worksheet chart (record 0270), added under its sheet.
    Chart,
    /// A worksheet feature (record 0320), with the rule type a CF element
    /// word implies.
    Feature(Kind, Option<&'static str>),
    /// A worksheet cell (record 0338), under a sheet, row or cell.
    Cell,
    /// A rich-text run (record 0338), under a cell.
    Run,
}

fn new_kind(name: &str, props: &crate::props::Props<'_>) -> Result<NewKind> {
    if name.eq_ignore_ascii_case("sheet") {
        Ok(NewKind::Sheet)
    } else if name.eq_ignore_ascii_case("chart") {
        Ok(NewKind::Chart)
    } else if name.eq_ignore_ascii_case("cell") {
        Ok(NewKind::Cell)
    } else if name.eq_ignore_ascii_case("run") || name.eq_ignore_ascii_case("r") {
        Ok(NewKind::Run)
    } else if let Some(implied) = super::feature_write::cf_element(name) {
        Ok(NewKind::Feature(Kind::Cf, Some(implied)))
    } else if name.eq_ignore_ascii_case("pagebreak") {
        // OfficeCLI's dispatcher: a column break when a column is named.
        let column = props.iter().any(|(key, _)| {
            key.as_str().eq_ignore_ascii_case("col") || key.as_str().eq_ignore_ascii_case("column")
        });
        Ok(NewKind::Feature(
            if column {
                Kind::ColBreak
            } else {
                Kind::RowBreak
            },
            None,
        ))
    } else if let Some(kind) = feature_kind(name) {
        Ok(NewKind::Feature(kind, None))
    } else {
        Err(Error::new(
            Code::UnsupportedElement,
            "this element kind cannot be added to an XLSX workbook",
        )
        .with_format(FileFormat::Xlsx)
        .with_suggestion("capabilities() lists the addable elements and the gaps"))
    }
}

/// The worksheet feature an element word names (`comment`, `note`,
/// `validation`, `image`, `textbox`).
fn feature_kind(name: &str) -> Option<Kind> {
    let canonical = match name.to_ascii_lowercase().as_str() {
        "img" => "picture",
        "textbox" => "shape",
        _ => crate::path::vocabulary::WORKBOOK.canonical(name)?,
    };
    let kind = <Kind as crate::path::navigate::ElementKind>::from_name(canonical)?;
    super::features::is_feature(kind).then_some(kind)
}

fn resolve_path<'f>(file: &'f Xlsx, path: &PathArg<'_>) -> Result<Node<'f>> {
    crate::path::resolve::get(&Nav::new(file), path)
}

fn sheet_position(node: &Node<'_>) -> Result<usize> {
    if node.element() != Kind::Sheet {
        return Err(Error::new(
            Code::UnsupportedType,
            "the XLSX structural verb requires a sheet tab",
        )
        .with_format(FileFormat::Xlsx));
    }
    Ok(node.sheet_position())
}

enum MoveDestination {
    Position(usize),
    Before(usize),
    After(usize),
}

fn move_destination(file: &Xlsx, placement: &Placement<'_>) -> Result<MoveDestination> {
    let last = file
        .0
        .owner
        .len()
        .checked_sub(1)
        .ok_or_else(|| not_found("the XLSX workbook has no sheet tabs"))?;
    match placement {
        Placement::Start => Ok(MoveDestination::Position(0)),
        Placement::End => Ok(MoveDestination::Position(last)),
        Placement::Position(position) => {
            if position.get() >= file.0.owner.len() {
                return Err(out_of_range(
                    "the XLSX sheet move position is outside the workbook",
                ));
            }
            Ok(MoveDestination::Position(position.get()))
        },
        Placement::Before(path) => {
            let anchor = sheet_position(&resolve_path(file, path)?)?;
            Ok(MoveDestination::Before(anchor))
        },
        Placement::After(path) => {
            let anchor = sheet_position(&resolve_path(file, path)?)?;
            Ok(MoveDestination::After(anchor))
        },
        Placement::Into(path, position) => {
            let parent = resolve_path(file, path)?;
            if parent.element() != Kind::Workbook {
                return Err(refusal(
                    "an XLSX sheet can only be moved within the workbook",
                ));
            }
            let position = position.map(Position::get).unwrap_or(last);
            if position >= file.0.owner.len() {
                return Err(out_of_range(
                    "the XLSX sheet move position is outside the workbook",
                ));
            }
            Ok(MoveDestination::Position(position))
        },
    }
}

fn add_at(
    edit: &mut xlsx::workbook::Edit,
    name: xlsx::sheet::Name,
    placement: Option<&Placement<'_>>,
    hidden: bool,
    file: &Xlsx,
) -> Result<usize> {
    let position = match placement {
        None | Some(Placement::End) => {
            let mut sheet = edit.add(name).map_err(native_error)?;
            set_new_sheet_visibility(&mut sheet, hidden);
            sheet.position()
        },
        Some(Placement::Start) => add_before_or_tail(edit, name, 0, hidden, file)?,
        Some(Placement::Position(position)) => {
            let position = position.get();
            if position > file.0.owner.len() {
                return Err(out_of_range(
                    "the XLSX sheet insertion position is outside the workbook",
                ));
            }
            if position == file.0.owner.len() {
                let mut sheet = edit.add(name).map_err(native_error)?;
                set_new_sheet_visibility(&mut sheet, hidden);
                sheet.position()
            } else {
                let mut sheet = edit
                    .add_before(name, position)
                    .map_err(native_error)?
                    .ok_or_else(|| not_found("the XLSX sheet insertion anchor disappeared"))?;
                set_new_sheet_visibility(&mut sheet, hidden);
                sheet.position()
            }
        },
        Some(Placement::Before(path)) => {
            let anchor = sheet_position(&resolve_path(file, path)?)?;
            let mut sheet = edit
                .add_before(name, anchor)
                .map_err(native_error)?
                .ok_or_else(|| not_found("the XLSX sheet insertion anchor disappeared"))?;
            set_new_sheet_visibility(&mut sheet, hidden);
            sheet.position()
        },
        Some(Placement::After(path)) => {
            let anchor = sheet_position(&resolve_path(file, path)?)?;
            let mut sheet = edit
                .add_after(name, anchor)
                .map_err(native_error)?
                .ok_or_else(|| not_found("the XLSX sheet insertion anchor disappeared"))?;
            set_new_sheet_visibility(&mut sheet, hidden);
            sheet.position()
        },
        Some(Placement::Into(path, position)) => {
            let parent = resolve_path(file, path)?;
            if parent.element() != Kind::Workbook {
                return Err(refusal(
                    "an XLSX sheet can only be added within the workbook",
                ));
            }
            let position = position.map(Position::get).unwrap_or(file.0.owner.len());
            if position > file.0.owner.len() {
                return Err(out_of_range(
                    "the XLSX sheet insertion position is outside the workbook",
                ));
            }
            if position == file.0.owner.len() {
                let mut sheet = edit.add(name).map_err(native_error)?;
                set_new_sheet_visibility(&mut sheet, hidden);
                sheet.position()
            } else {
                let mut sheet = edit
                    .add_before(name, position)
                    .map_err(native_error)?
                    .ok_or_else(|| not_found("the XLSX sheet insertion anchor disappeared"))?;
                set_new_sheet_visibility(&mut sheet, hidden);
                sheet.position()
            }
        },
    };
    Ok(position)
}

fn add_before_or_tail(
    edit: &mut xlsx::workbook::Edit,
    name: xlsx::sheet::Name,
    position: usize,
    hidden: bool,
    file: &Xlsx,
) -> Result<usize> {
    if file.0.owner.len() == 0 {
        let mut sheet = edit.add(name).map_err(native_error)?;
        set_new_sheet_visibility(&mut sheet, hidden);
        return Ok(sheet.position());
    }
    let mut sheet = edit
        .add_before(name, position)
        .map_err(native_error)?
        .ok_or_else(|| not_found("the XLSX sheet insertion anchor disappeared"))?;
    set_new_sheet_visibility(&mut sheet, hidden);
    Ok(sheet.position())
}

fn set_new_sheet_visibility(sheet: &mut xlsx::workbook::NewSheet<'_>, hidden: bool) {
    if hidden {
        sheet.hide();
    }
}

fn sheet_path(position: usize) -> Path<'static> {
    Path::from_steps([Step::new("sheet").with(Predicate::Index(Position::new(position)))])
}

fn native_error(error: xlsx::Error) -> Error {
    super::sheet_write::sheet_error(error)
}

const REMOVE_ENTRY_OVERHEAD: usize = 64;

fn check_removal_limits(count: usize, limits: &SemanticLimits) -> Result<()> {
    if count > limits.max_expanded_targets() {
        return Err(Error::from(TransactionError::Limit {
            kind: LimitKind::ExpandedTargets,
            observed: count,
            maximum: limits.max_expanded_targets(),
        })
        .with_format(FileFormat::Xlsx));
    }
    let staged = count.saturating_mul(REMOVE_ENTRY_OVERHEAD);
    if staged > limits.max_staged_bytes() {
        return Err(Error::from(TransactionError::Limit {
            kind: LimitKind::StagedBytes,
            observed: staged,
            maximum: limits.max_staged_bytes(),
        })
        .with_format(FileFormat::Xlsx));
    }
    Ok(())
}

fn unsupported_removal_target() -> Error {
    Error::new(
        Code::UnsupportedElement,
        "the XLSX sheet removal target must be a sheet",
    )
    .with_format(FileFormat::Xlsx)
}

fn refusal(message: &'static str) -> Error {
    Error::new(Code::UnsupportedCapability, message).with_format(FileFormat::Xlsx)
}

fn out_of_range(message: &'static str) -> Error {
    Error::new(Code::OutOfRange, message).with_format(FileFormat::Xlsx)
}

fn not_found(message: &'static str) -> Error {
    Error::new(Code::NotFound, message).with_format(FileFormat::Xlsx)
}
