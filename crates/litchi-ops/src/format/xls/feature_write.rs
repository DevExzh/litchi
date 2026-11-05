//! Add, set and remove of XLS worksheet features (record 0327).
//!
//! Every verb admits its properties through the kind's registry table in
//! `features` (the XLSX vocabulary of record 0320), builds the leaf value
//! and publishes one `litchi_xls::sheet_features::edit` commit: notes,
//! hyperlinks, pictures, shapes, conditional formats, validations, the
//! auto-filter and page breaks are record-level edits of the worksheet
//! substream (and of the drawing group for drawings), so every other record
//! keeps its bytes.
//!
//! Formulas (CF values and rules, validation formulas) are text without a
//! leading `=`, compiled by the BIFF8 tokenizer: same-sheet references,
//! numbers, strings, operators and built-in functions. Colors of a
//! conditional format's formatting are palette entries of the workbook (the
//! BIFF8 `DXFN` stores palette indexes); drawing and data-bar colors are any
//! RGB color. Drawings are anchored to cells, so their geometry is in cells.

use litchi_core::selector::{Path, Predicate, Step};
use litchi_core::{FileFormat, Position};
use litchi_xls::sheet_features::{self as sf, RuleType};
use sf::edit::{
    Anchor, DrawingChange, Format, FormatRule, Hyperlink as LinkSpec, PageBreak,
    Picture as PictureSpec, Shape as ShapeSpec, ShapeGeometry, Style as StyleSpec,
    Threshold as Cfvo, Validation as ValidationSpec,
};
use std::mem::size_of;

use super::features::{
    self, AutofilterKey, BreakKey, CfKey, CommentKey, Data, Feature, HyperlinkKey, PictureKey,
    ShapeKey, ValidationKey,
};
use super::node::{Item, Node};
use super::spec::Kind;
use super::write::{StepPatch, WriteCommit};
use super::{Nav, Xls};
use crate::error::{Code, Error, Result, refusal_code};
use crate::new_node::NewNode;
use crate::placement::Placement as NodePlacement;
use crate::props::Props;
use crate::registry::{Operation, Table};
use crate::target::TargetArg;
use crate::transaction::SemanticLimits;
use crate::value::{Color, ColorContext, Scalar};

fn error(code: Code, message: impl Into<String>) -> Error {
    Error::new(code, message.into()).with_format(FileFormat::Xls)
}

fn invalid(key: &'static str, message: impl Into<String>) -> Error {
    error(Code::InvalidValue, message).with_property(key)
}

/// A leaf refusal as its unified code.
fn native(leaf: sf::edit::Error) -> Error {
    use sf::edit::Error as Leaf;
    match leaf {
        Leaf::Refused(refusal) => error(refusal_code(refusal.kind()), refusal.reason()),
        Leaf::Limit {
            observed, maximum, ..
        } => error(
            Code::LimitExceeded,
            format!("the worksheet drawing holds at most {maximum} shapes ({observed} requested)"),
        ),
        Leaf::PatchConflict => error(
            Code::InvalidInput,
            "a worksheet-feature patch met another package",
        ),
        Leaf::Signed => super::signature::refusal(FileFormat::Xls),
        Leaf::Package(error) => Error::from(error).with_format(FileFormat::Xls),
        Leaf::Records(error) => Error::from(error).with_format(FileFormat::Xls),
        _ => error(Code::InternalError, "the worksheet-feature edit failed"),
    }
}

// ---------------------------------------------------------------------------
// Publication

impl Xls {
    /// Publish one leaf commit as a unified write.
    fn publish_features(
        &self,
        commit: sf::edit::Commit,
        limits: &SemanticLimits,
    ) -> Result<WriteCommit> {
        // A signed setter may still prove an exact native identity patch.
        // Changed feature output is rejected below, but the strict guard here
        // would prevent same-value set/remove requests from reaching that
        // proof.
        self.check_source_for_write()?;
        let changed = commit.diagnostics().changed();
        let staged_snapshot_bytes = commit.patch().retained_bytes();
        let (package, patch) = commit.into_parts();
        if changed && self.signature() == super::Signature::Signed {
            return Err(super::signature::refusal(FileFormat::Xls));
        }
        let (package, patch, snapshot_bytes) =
            if !changed && self.signature() == super::Signature::Signed {
                // The feature probe is staged on an unsigned private copy.
                // Rebuild the exact identity patch over the retained signed
                // package before publishing the no-op.
                let original = self.original_package_for_noop()?;
                let minimum = self
                    .retained_source_len()
                    .saturating_mul(2)
                    .saturating_add(size_of::<litchi_xls::Patch>())
                    .saturating_add(size_of::<litchi_xls::sheet_features::edit::Edit>());
                if minimum > limits.max_snapshot_bytes() {
                    return Err(error(
                        Code::LimitExceeded,
                        "the worksheet-feature identity patch exceeds the snapshot budget",
                    ));
                }
                let empty = original
                    .edit_sheet_features()
                    .map_err(native)?
                    .commit()
                    .map_err(native)?;
                let canonical_snapshot_bytes = empty.patch().retained_bytes();
                let (package, patch) = empty.into_parts();
                (
                    package,
                    patch,
                    staged_snapshot_bytes.max(canonical_snapshot_bytes),
                )
            } else {
                (package, patch, staged_snapshot_bytes)
            };
        if snapshot_bytes > limits.max_snapshot_bytes() {
            return Err(error(
                Code::LimitExceeded,
                "the worksheet-feature patch exceeds the snapshot budget",
            ));
        }
        if limits.max_operations() < 1 {
            return Err(error(
                Code::LimitExceeded,
                "the operation budget admits no worksheet-feature write",
            ));
        }
        let after = if changed {
            self.with_package(package)?
        } else {
            self.clone()
        };
        self.check_source_for_write()?;
        Ok(WriteCommit {
            before: self.clone(),
            after,
            patch: StepPatch::Features(patch),
            // One feature command remains one authored operation even when
            // its native patch is the canonical empty identity.
            operations: 1,
            snapshot_bytes,
        })
    }

    fn feature_edit(&self, maximum: usize) -> Result<sf::edit::Edit> {
        self.preflight_snapshot_bytes(
            maximum,
            size_of::<litchi_xls::sheet_features::edit::Edit>(),
        )?;
        self.editable_package_for_write_with_snapshot_limit(maximum)?
            .edit_sheet_features()
            .map_err(native)
    }
}

/// Run one leaf edit and publish it.
macro_rules! publish {
    ($file:expr, $limits:expr, |$edit:ident| $body:expr) => {{
        let mut $edit = $file.feature_edit($limits.max_snapshot_bytes())?;
        let value = { $body };
        let commit = $edit.commit().map_err(native)?;
        ($file.publish_features(commit, $limits)?, value)
    }};
}

// ---------------------------------------------------------------------------
// Property admission

/// An admitted property value, detached from the request.
#[derive(Debug, Clone)]
enum Arg {
    Null,
    Text(String),
    Bool(bool),
    Int(i64),
    Number(f64),
    Length(crate::value::Length),
    Color(Color),
}

impl Arg {
    fn of(scalar: Option<Scalar<'_>>) -> Self {
        match scalar {
            None => Self::Null,
            Some(Scalar::Text(value)) => Self::Text(value.to_owned()),
            Some(Scalar::Token(value)) => Self::Text(value.to_owned()),
            Some(Scalar::Bool(value)) => Self::Bool(value),
            Some(Scalar::Int(value)) => Self::Int(value),
            Some(Scalar::Number(value)) => Self::Number(value),
            Some(Scalar::Length(value)) => Self::Length(value),
            Some(Scalar::Color(value)) => Self::Color(value),
            Some(Scalar::List(items)) => Self::Text(items.join(", ")),
        }
    }

    /// The value as text; numbers and flags in their plain spelling.
    fn text(&self) -> Option<String> {
        match self {
            Self::Null => None,
            Self::Text(value) => Some(value.clone()),
            Self::Bool(value) => Some(value.to_string()),
            Self::Int(value) => Some(value.to_string()),
            Self::Number(value) => Some(crate::value::DotNetDouble(*value).to_string()),
            Self::Length(value) => Some(value.emu().to_string()),
            Self::Color(value) => Some(value.to_string()),
        }
    }

    fn flag(&self) -> bool {
        matches!(self, Self::Bool(true))
    }

    fn int(&self, key: &'static str) -> Result<i64> {
        match self {
            Self::Int(value) => Ok(*value),
            Self::Text(value) => value
                .trim()
                .parse()
                .map_err(|_| invalid(key, "the value is not an integer")),
            _ => Err(invalid(key, "the value is not an integer")),
        }
    }

    fn number(&self, key: &'static str) -> Result<f64> {
        match self {
            Self::Int(value) =>
            {
                #[allow(clippy::cast_precision_loss, reason = "angles are small")]
                Ok(*value as f64)
            },
            Self::Number(value) => Ok(*value),
            Self::Text(value) => value
                .trim()
                .parse()
                .map_err(|_| invalid(key, "the value is not a number")),
            _ => Err(invalid(key, "the value is not a number")),
        }
    }
}

/// Admit `props` through `table` for `operation`: each property once.
fn admit<S: Copy + PartialEq, const N: usize, const A: usize>(
    table: &Table<S, N, A>,
    operation: Operation,
    props: &Props<'_>,
) -> Result<Vec<(S, &'static str, Arg)>> {
    let mut admitted: Vec<(S, &'static str, Arg)> = Vec::with_capacity(props.len());
    for (key, value) in props.iter() {
        let (_, entry) = table
            .resolve(key.as_str())
            .map_err(|error| error.with_property(key.as_str().to_owned()))?;
        let name = entry.spec.name;
        let (setter, arg) = entry.prepare(operation, value, |setter, scalar| {
            Ok((setter, Arg::of(scalar)))
        })?;
        if admitted.iter().any(|(held, _, _)| *held == setter) {
            return Err(error(
                Code::InvalidArgument,
                "a property was assigned more than once",
            )
            .with_property(name));
        }
        admitted.push((setter, name, arg));
    }
    Ok(admitted)
}

fn feature_path(sheet: usize, step: &str, index: Option<usize>) -> Path<'static> {
    let mut last = Step::new(step.to_owned());
    if let Some(index) = index {
        last = last.with(Predicate::Index(Position::new(index)));
    }
    Path::from_steps([
        Step::new("sheet").with(Predicate::Index(Position::new(sheet))),
        last,
    ])
}

/// A zero-based cell `(row, column)` of the BIFF8 grid.
type Cell = (u16, u16);

/// Parse an A1 cell (dollar signs ignored) inside A1:IV65536.
fn parse_cell(text: &str) -> Option<Cell> {
    let cell = litchi_core::selector::CellRef::parse(&text.trim().replace('$', ""))?;
    let row = u16::try_from(cell.row()).ok()?;
    let column = u16::try_from(cell.column())
        .ok()
        .filter(|column| *column <= 255)?;
    Some((row, column))
}

/// Parse an A1 cell or range (dollar signs ignored).
fn parse_area(text: &str) -> Option<sf::Area> {
    let text = text.trim();
    match text.split_once(':') {
        Some((first, last)) => {
            let (first, last) = (parse_cell(first)?, parse_cell(last)?);
            (first.0 <= last.0 && first.1 <= last.1)
                .then(|| sf::Area::new(first.0, first.1, last.0, last.1))
        },
        None => parse_cell(text).map(|cell| sf::Area::cell(cell.0, cell.1)),
    }
}

/// Space-separated A1 ranges.
fn parse_ranges(key: &'static str, value: &str) -> Result<Vec<sf::Area>> {
    let ranges = value
        .split_whitespace()
        .map(|range| {
            parse_area(range)
                .ok_or_else(|| invalid(key, "the range is not an A1 range inside A1:IV65536"))
        })
        .collect::<Result<Vec<_>>>()?;
    if ranges.is_empty() {
        return Err(invalid(key, "the range is empty"));
    }
    Ok(ranges)
}

/// Where a feature is added: its worksheet, and the cell or range the
/// parent path named.
struct Parent {
    sheet: usize,
    reference: Option<String>,
}

fn parent(node: &Node<'_>) -> Result<Parent> {
    let sheet = node.sheet_position();
    match node.item() {
        Item::Sheet(_) => Ok(Parent {
            sheet,
            reference: None,
        }),
        Item::Cell(cell) => Ok(Parent {
            sheet,
            reference: Some(cell.label.as_str().to_owned()),
        }),
        Item::Range(range) => Ok(Parent {
            sheet,
            reference: Some(range.label.as_str().to_owned()),
        }),
        _ => Err(unsupported_parent()),
    }
}

fn unsupported_parent() -> Error {
    error(
        Code::UnsupportedType,
        "worksheet features are added under a worksheet, a cell or a range",
    )
}

fn missing(key: &'static str, message: &'static str) -> Error {
    error(Code::MissingProperty, message).with_property(key)
}

/// The BIFF8 tokens of a formula (text without a leading `=`).
fn formula_tokens(key: &'static str, formula: &str) -> Result<Vec<u8>> {
    let formula = formula.trim();
    if formula.is_empty() {
        return Err(invalid(key, "the formula is empty"));
    }
    sf::edit::compile_formula(formula).map_err(|_| {
        invalid(key, "the formula cannot be encoded as BIFF8 tokens").with_suggestion(
            "use same-sheet references, numbers, text, operators and built-in functions",
        )
    })
}

// ---------------------------------------------------------------------------
// Dispatch

impl Xls {
    /// The features of `kind` on the worksheet at `sheet`, read again.
    fn features_of(&self, sheet: usize, kind: Kind) -> Result<Vec<Feature>> {
        let handle = super::owner::SheetHandle::at(self, sheet)?
            .ok_or_else(|| error(Code::PathNotFound, "the worksheet is missing"))?;
        features::list_handle(self, &handle, kind)
    }

    /// Add a worksheet feature under `parent`; `implied` is the rule type a
    /// CF element word (`databar`, `cellis`) names.
    pub(super) fn add_feature(
        &self,
        parent_node: &Node<'_>,
        kind: Kind,
        implied: Option<&'static str>,
        node: &NewNode<'_>,
        limits: &SemanticLimits,
    ) -> Result<(WriteCommit, Path<'static>)> {
        if !matches!(node.placement(), None | Some(NodePlacement::End)) {
            return Err(error(
                Code::UnsupportedCapability,
                "worksheet features are added at the end of their kind",
            )
            .with_suggestion("omit placement or use Placement::End"));
        }
        let parent = parent(parent_node)?;
        if !super::owner::SheetHandle::at(self, parent.sheet)?
            .is_some_and(|handle| handle.is_grid())
        {
            return Err(error(
                Code::UnsupportedType,
                "worksheet features belong to worksheets",
            ));
        }
        let props = node.props();
        match kind {
            Kind::Comment => self.add_comment(&parent, props, limits),
            Kind::Hyperlink => self.add_hyperlink(&parent, props, limits),
            Kind::Picture => self.add_picture(&parent, props, limits),
            Kind::Shape => self.add_shape(&parent, props, limits),
            Kind::Cf => self.add_cf(&parent, implied, props, limits),
            Kind::DataValidation => self.add_validation(&parent, props, limits),
            Kind::Autofilter => self.add_autofilter(&parent, props, limits),
            Kind::RowBreak | Kind::ColBreak => self.add_break(&parent, kind, props, limits),
            _ => Err(error(
                Code::UnsupportedElement,
                "this element kind cannot be added to an XLS worksheet",
            )),
        }
    }

    /// The feature a path names; `None` for other targets.
    fn feature_target<'f>(&'f self, target: &TargetArg<'_>) -> Result<Option<Node<'f>>> {
        let Some(path) = target.path() else {
            return Ok(None);
        };
        let Some(last) = path.path().steps().last() else {
            return Ok(None);
        };
        let Some(kind) = crate::path::vocabulary::WORKBOOK
            .canonical(last.name())
            .and_then(<Kind as crate::path::navigate::ElementKind>::from_name)
        else {
            return Ok(None);
        };
        if !features::is_feature(kind) {
            return Ok(None);
        }
        let node = crate::path::resolve::get(&Nav::new(self), path)?;
        Ok(matches!(node.item(), Item::Feature(_)).then_some(node))
    }

    /// A `set` on a feature path; `None` for other targets.
    pub(super) fn try_set_feature(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<WriteCommit>> {
        let Some(node) = self.feature_target(target)? else {
            return Ok(None);
        };
        let Item::Feature(feature) = node.item() else {
            return Ok(None);
        };
        let sheet = node.sheet_position();
        Ok(Some(match feature.kind {
            Kind::Comment => self.set_comment(sheet, feature, props, limits)?,
            Kind::Hyperlink => self.set_hyperlink(sheet, feature, props, limits)?,
            Kind::Picture | Kind::Shape => self.set_drawing(sheet, feature, props, limits)?,
            Kind::Cf => self.set_cf(sheet, feature, props, limits)?,
            Kind::DataValidation => self.set_validation(sheet, feature, props, limits)?,
            Kind::Autofilter => self.set_autofilter(sheet, props, limits)?,
            _ => self.set_break(sheet, feature, props, limits)?,
        }))
    }

    /// Remove the feature a path names; `None` for other targets.
    pub(super) fn try_remove_feature(
        &self,
        target: &TargetArg<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<WriteCommit>> {
        let Some(node) = self.feature_target(target)? else {
            return Ok(None);
        };
        let Item::Feature(feature) = node.item() else {
            return Ok(None);
        };
        let sheet = node.sheet_position();
        let index = feature.index as usize;
        let (commit, ()) = match &feature.data {
            Data::Comment(note) => {
                let (row, col) = note_cell(note)?;
                publish!(self, limits, |edit| edit
                    .remove_note(sheet, row, col)
                    .map_err(native)?)
            },
            Data::Hyperlink(_) => {
                publish!(self, limits, |edit| edit
                    .remove_hyperlink(sheet, index)
                    .map_err(native)?)
            },
            Data::Drawing(_) => {
                let picture = feature.kind == Kind::Picture;
                publish!(self, limits, |edit| edit
                    .remove_drawing(sheet, picture, index)
                    .map_err(native)?)
            },
            Data::Cf(_) => {
                publish!(self, limits, |edit| edit
                    .remove_format(sheet, index)
                    .map_err(native)?)
            },
            Data::Validation(_) => {
                publish!(self, limits, |edit| edit
                    .remove_validation(sheet, index)
                    .map_err(native)?)
            },
            Data::Autofilter(_) => {
                publish!(self, limits, |edit| edit
                    .set_auto_filter(sheet, None)
                    .map_err(native)?)
            },
            Data::Break(_) => {
                let mut breaks = self.breaks(sheet, feature.kind)?;
                if index >= breaks.len() {
                    return Err(error(Code::PathNotFound, "the page break is missing"));
                }
                breaks.remove(index);
                return self
                    .publish_breaks(sheet, feature.kind, &breaks, limits)
                    .map(Some);
            },
        };
        Ok(Some(commit))
    }
}

// ---------------------------------------------------------------------------
// Notes

fn check_cell(key: &'static str, reference: &str) -> Result<Cell> {
    parse_cell(reference)
        .ok_or_else(|| invalid(key, "the reference is not one A1 cell inside A1:IV65536"))
}

fn note_cell(note: &features::Note) -> Result<Cell> {
    Ok((
        u16::try_from(note.row)
            .map_err(|_| error(Code::InternalError, "a note row exceeds 16 bits"))?,
        u16::try_from(note.col)
            .map_err(|_| error(Code::InternalError, "a note column exceeds 16 bits"))?,
    ))
}

impl Xls {
    fn add_comment(
        &self,
        parent: &Parent,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<(WriteCommit, Path<'static>)> {
        let mut reference = None;
        let mut author = None;
        let mut text = None;
        for (key, name, arg) in admit(&features::COMMENT_TABLE, Operation::Add, props)? {
            match key {
                CommentKey::Ref => {
                    reference = Some(check_cell(name, &arg.text().unwrap_or_default())?);
                },
                CommentKey::Author => author = arg.text(),
                CommentKey::Text => text = arg.text(),
                CommentKey::AnchoredTo => {},
            }
        }
        let (row, col) = match (reference, &parent.reference) {
            (Some(cell), _) => cell,
            (None, Some(cell)) => check_cell("ref", cell)?,
            (None, None) => return Err(missing("ref", "a note needs a ref")),
        };
        let author = author.unwrap_or_else(|| "Author".to_owned());
        let text = text.unwrap_or_default();
        let (commit, index) = publish!(self, limits, |edit| edit
            .add_note(parent.sheet, row, col, &author, &text)
            .map_err(native)?);
        // Notes are numbered row-major, as `features::list` numbers them.
        let Some(index) = index else {
            return Err(error(Code::InvalidArgument, "the cell already has a note")
                .with_property("ref")
                .with_suggestion("set the existing comment instead"));
        };
        Ok((commit, feature_path(parent.sheet, "comment", Some(index))))
    }

    fn set_comment(
        &self,
        sheet: usize,
        feature: &Feature,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<WriteCommit> {
        let Data::Comment(note) = &feature.data else {
            return Err(error(Code::InternalError, "a note node holds no note"));
        };
        let cell = note_cell(note)?;
        let mut target = cell;
        let mut author = note.author.clone();
        let mut text = note.text.clone();
        for (key, name, arg) in admit(&features::COMMENT_TABLE, Operation::Set, props)? {
            match key {
                CommentKey::Ref => target = check_cell(name, &arg.text().unwrap_or_default())?,
                CommentKey::Author => author = arg.text().unwrap_or_default(),
                CommentKey::Text => text = arg.text().unwrap_or_default(),
                CommentKey::AnchoredTo => {},
            }
        }
        let (commit, ()) = publish!(self, limits, |edit| {
            edit.set_note(sheet, cell, target, &author, &text)
                .map_err(native)?;
        });
        Ok(commit)
    }
}

// ---------------------------------------------------------------------------
// Hyperlinks

/// Unsafe link schemes OfficeCLI refuses.
fn check_url(key: &'static str, url: &str) -> Result<()> {
    let lower = url.trim().to_ascii_lowercase();
    if ["javascript:", "vbscript:", "data:", "file:"]
        .iter()
        .any(|scheme| lower.starts_with(scheme))
    {
        return Err(invalid(key, "the link scheme is not allowed"));
    }
    Ok(())
}

/// A link value as (location, external target): `#Sheet!A1` and `Sheet!A1`
/// are internal, anything else an external target.
fn link_target(key: &'static str, value: &str) -> Result<(Option<String>, Option<String>)> {
    if let Some(location) = value.strip_prefix('#') {
        return Ok((Some(location.to_owned()), None));
    }
    if !value.contains("://") && !value.contains(':') && value.contains('!') {
        return Ok((Some(value.to_owned()), None));
    }
    check_url(key, value)?;
    Ok((None, Some(value.to_owned())))
}

fn hyperlink(
    reference: &str,
    location: Option<String>,
    url: Option<String>,
    display: Option<String>,
    tooltip: Option<String>,
) -> Result<LinkSpec> {
    let area =
        parse_area(reference).ok_or_else(|| invalid("ref", "the reference is not an A1 range"))?;
    if location.is_none() && url.is_none() {
        return Err(missing("url", "a hyperlink needs a url or a location"));
    }
    if tooltip
        .as_ref()
        .is_some_and(|tooltip| tooltip.encode_utf16().count() > 255)
    {
        return Err(invalid("tooltip", "a tooltip holds at most 255 characters"));
    }
    Ok(LinkSpec {
        area,
        url,
        location,
        display,
        tooltip,
    })
}

impl Xls {
    fn add_hyperlink(
        &self,
        parent: &Parent,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<(WriteCommit, Path<'static>)> {
        let mut reference = parent.reference.clone();
        let mut location = None;
        let mut url = None;
        let mut display = None;
        let mut tooltip = None;
        for (key, name, arg) in admit(&features::HYPERLINK_TABLE, Operation::Add, props)? {
            match key {
                HyperlinkKey::Ref => reference = arg.text(),
                HyperlinkKey::Url => {
                    if let Some(value) = arg.text() {
                        let (internal, external) = link_target(name, &value)?;
                        location = internal.or(location);
                        url = external;
                    }
                },
                HyperlinkKey::Location => location = arg.text(),
                HyperlinkKey::Display => display = arg.text(),
                HyperlinkKey::Tooltip => tooltip = arg.text(),
            }
        }
        let reference = reference.ok_or_else(|| missing("ref", "a hyperlink needs a ref"))?;
        let link = hyperlink(&reference, location, url, display, tooltip)?;
        // As XLSX's `put_hyperlink`: a link on the same cells is replaced.
        let label = features::area_label(link.area);
        let existing = self
            .features_of(parent.sheet, Kind::Hyperlink)?
            .into_iter()
            .position(|feature| match &feature.data {
                Data::Hyperlink(held) => features::link_reference(held) == label,
                _ => false,
            });
        let (commit, index) = publish!(self, limits, |edit| match existing {
            Some(index) => {
                edit.set_hyperlink(parent.sheet, index, &link)
                    .map_err(native)?;
                index
            },
            None => edit.add_hyperlink(parent.sheet, &link).map_err(native)?,
        });
        Ok((commit, feature_path(parent.sheet, "hyperlink", Some(index))))
    }

    fn set_hyperlink(
        &self,
        sheet: usize,
        feature: &Feature,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<WriteCommit> {
        let Data::Hyperlink(link) = &feature.data else {
            return Err(error(
                Code::InternalError,
                "a hyperlink node holds no hyperlink",
            ));
        };
        let mut reference = features::link_reference(link);
        let mut location = link.location().map(str::to_owned);
        let mut url = features::link_url(link).map(str::to_owned);
        let mut display = link.display_name().map(str::to_owned);
        let mut tooltip = link.tooltip().map(str::to_owned);
        for (key, name, arg) in admit(&features::HYPERLINK_TABLE, Operation::Set, props)? {
            match key {
                HyperlinkKey::Ref => reference = arg.text().unwrap_or_default(),
                HyperlinkKey::Url => match arg.text() {
                    Some(value) => {
                        let (internal, external) = link_target(name, &value)?;
                        if internal.is_some() {
                            location = internal;
                        }
                        url = external;
                    },
                    None => url = None,
                },
                HyperlinkKey::Location => location = arg.text(),
                HyperlinkKey::Display => display = arg.text(),
                HyperlinkKey::Tooltip => tooltip = arg.text(),
            }
        }
        let replacement = hyperlink(&reference, location, url, display, tooltip)?;
        let index = feature.index as usize;
        let (commit, ()) = publish!(self, limits, |edit| edit
            .set_hyperlink(sheet, index, &replacement)
            .map_err(native)?);
        Ok(commit)
    }
}

// ---------------------------------------------------------------------------
// Pictures and shapes

/// The largest image a picture embeds (the BIFF8 drawing group continues
/// past 8,224 bytes, so this bounds memory, not the format).
const MAX_IMAGE_BYTES: usize = 64 * 1024 * 1024;

/// A picture source: a base64 `data:` URI or a local file path.
fn read_source(src: &str) -> Result<Vec<u8>> {
    if let Some(rest) = src.strip_prefix("data:") {
        let (header, payload) = rest
            .split_once(',')
            .ok_or_else(|| invalid("src", "a data URI needs a comma before its payload"))?;
        if !header.ends_with(";base64") {
            return Err(error(
                Code::UnsupportedCapability,
                "only base64 data URIs are accepted",
            )
            .with_property("src"));
        }
        if payload.len() / 4 * 3 > MAX_IMAGE_BYTES {
            return Err(invalid("src", "the image is too large"));
        }
        use base64::Engine as _;
        return base64::engine::general_purpose::STANDARD
            .decode(payload.trim())
            .map_err(|_| invalid("src", "the data URI payload is not base64"));
    }
    let length = std::fs::metadata(src)
        .map_err(|source| Error::from(source).with_property("src"))?
        .len();
    if length > MAX_IMAGE_BYTES as u64 {
        return Err(invalid("src", "the image is too large"));
    }
    std::fs::read(src).map_err(|source| Error::from(source).with_property("src"))
}

/// A geometry value in cells.
fn cells(key: &'static str, arg: &Arg) -> Result<u16> {
    let value = match arg {
        Arg::Int(value) => *value,
        Arg::Text(value) if value.trim().bytes().all(|byte| byte.is_ascii_digit()) => value
            .trim()
            .parse()
            .map_err(|_| invalid(key, "the cell count is out of range"))?,
        _ => {
            return Err(invalid(
                key,
                "an XLS drawing is anchored to cells: give a cell count",
            )
            .with_suggestion("use cell counts or anchor=B2:E6"));
        },
    };
    u16::try_from(value).map_err(|_| invalid(key, "the cell count is out of range"))
}

/// The geometry properties of a picture or shape request.
#[derive(Default)]
struct Geometry {
    x: Option<u16>,
    y: Option<u16>,
    width: Option<u16>,
    height: Option<u16>,
    anchor: Option<String>,
    mode: Option<String>,
}

impl Geometry {
    fn is_empty(&self) -> bool {
        self.x.is_none()
            && self.y.is_none()
            && self.width.is_none()
            && self.height.is_none()
            && self.anchor.is_none()
            && self.mode.is_none()
    }

    /// The anchor, from `current` (a set) or defaults `[x, y, width,
    /// height]` (an add).
    fn anchor(
        &self,
        current: Option<litchi_xls::SheetAnchor>,
        defaults: [u16; 4],
    ) -> Result<Anchor> {
        let behavior = match self.mode.as_deref().map(str::to_ascii_lowercase).as_deref() {
            Some("onecell") => litchi_xls::AnchorBehavior::MoveWithoutResize,
            Some("absolute") => litchi_xls::AnchorBehavior::Absolute,
            Some("twocell") => litchi_xls::AnchorBehavior::MoveAndResize,
            Some(_) => {
                return Err(invalid(
                    "anchorMode",
                    "anchorMode is oneCell, twoCell or absolute",
                ));
            },
            None => current.map_or(litchi_xls::AnchorBehavior::MoveAndResize, |anchor| {
                anchor.behavior()
            }),
        };
        if let Some(anchor) = &self.anchor {
            let anchor = anchor.trim().replace('$', "");
            let area = if anchor.contains(':') {
                parse_area(&anchor)
                    .ok_or_else(|| invalid("anchor", "the anchor is not an A1 range"))?
            } else {
                // A single cell spans four columns and two rows.
                let (row, column) = parse_cell(&anchor)
                    .ok_or_else(|| invalid("anchor", "the anchor is not an A1 range"))?;
                sf::Area::new(row, column, row + 1, column + 3)
            };
            return Ok(Anchor {
                first_row: area.first_row,
                first_column: area.first_column,
                last_row: area.last_row.saturating_add(1),
                last_column: (area.last_column + 1).min(255),
                behavior,
            });
        }
        let (x, y, width, height) = match current {
            Some(anchor) => (
                anchor.top_left().column(),
                anchor.top_left().row(),
                anchor
                    .bottom_right()
                    .column()
                    .saturating_sub(anchor.top_left().column()),
                anchor
                    .bottom_right()
                    .row()
                    .saturating_sub(anchor.top_left().row()),
            ),
            None => (defaults[0], defaults[1], defaults[2], defaults[3]),
        };
        let x = self.x.unwrap_or(x);
        let y = self.y.unwrap_or(y);
        let width = self.width.unwrap_or(width);
        let height = self.height.unwrap_or(height);
        if width == 0 || height == 0 {
            return Err(invalid("width", "a size must be positive"));
        }
        let last_column = x.checked_add(width).filter(|column| *column <= 255);
        let last_row = y.checked_add(height);
        match (last_column, last_row) {
            (Some(last_column), Some(last_row)) => Ok(Anchor {
                first_row: y,
                first_column: x,
                last_row,
                last_column,
                behavior,
            }),
            _ => Err(invalid("width", "the drawing ends outside A1:IV65536")),
        }
    }
}

fn flip_value(key: &'static str, value: &str) -> Result<(bool, bool)> {
    Ok(match value.trim().to_ascii_lowercase().as_str() {
        "h" | "horizontal" => (true, false),
        "v" | "vertical" => (false, true),
        "both" | "hv" | "vh" => (true, true),
        "none" | "" => (false, false),
        _ => return Err(invalid(key, "flip is h, v, both or none")),
    })
}

/// A plain RGB color as `0xRRGGBB`.
fn plain_rgb(key: &'static str, color: &Color) -> Result<u32> {
    match color.base() {
        crate::value::Base::Rgb(value) if color.transforms().is_empty() => {
            Ok(u32::from_be_bytes([
                0,
                value.red(),
                value.green(),
                value.blue(),
            ]))
        },
        _ => Err(invalid(key, "the color must be an RGB color")),
    }
}

fn parse_color(key: &'static str, value: &str, context: ColorContext) -> Result<u32> {
    let color = Color::parse(value.trim(), context)
        .map_err(|_| invalid(key, "the value is not a color"))?;
    plain_rgb(key, &color)
}

fn arg_color(key: &'static str, arg: &Arg, context: ColorContext) -> Result<u32> {
    match arg {
        Arg::Color(color) => plain_rgb(key, color),
        Arg::Text(value) => parse_color(key, value, context),
        _ => Err(invalid(key, "the value is not a color")),
    }
}

/// The geometry of an OfficeCLI preset name.
fn shape_geometry(key: &'static str, value: &str) -> Result<ShapeGeometry> {
    Ok(match value.trim().to_ascii_lowercase().as_str() {
        "rect" | "rectangle" => ShapeGeometry::Rectangle,
        "roundrect" | "roundedrectangle" => ShapeGeometry::RoundedRectangle,
        "ellipse" | "oval" => ShapeGeometry::Ellipse,
        "line" => ShapeGeometry::Line,
        "textbox" => ShapeGeometry::TextBox,
        _ => {
            return Err(invalid(
                key,
                "the legacy drawing writes rect, roundRect, ellipse, line and textbox",
            )
            .with_suggestion("use rect, roundRect, ellipse, line or textbox"));
        },
    })
}

impl Xls {
    fn add_picture(
        &self,
        parent: &Parent,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<(WriteCommit, Path<'static>)> {
        let mut src = None;
        let mut name = None;
        let mut description = None;
        let mut rotation = 0.0;
        let (mut flip_h, mut flip_v) = (false, false);
        let mut geometry = Geometry::default();
        for (key, name_key, arg) in admit(&features::PICTURE_TABLE, Operation::Add, props)? {
            match key {
                PictureKey::Src => src = arg.text(),
                PictureKey::Name => name = arg.text(),
                PictureKey::Alt => description = arg.text(),
                PictureKey::Rotation => rotation = arg.number(name_key)?,
                PictureKey::Flip => {
                    (flip_h, flip_v) = flip_value(name_key, &arg.text().unwrap_or_default())?;
                },
                PictureKey::FlipH => flip_h = arg.flag(),
                PictureKey::FlipV => flip_v = arg.flag(),
                PictureKey::FlipBoth => (flip_h, flip_v) = (arg.flag(), arg.flag()),
                PictureKey::X => geometry.x = Some(cells(name_key, &arg)?),
                PictureKey::Y => geometry.y = Some(cells(name_key, &arg)?),
                PictureKey::Width => geometry.width = Some(cells(name_key, &arg)?),
                PictureKey::Height => geometry.height = Some(cells(name_key, &arg)?),
                PictureKey::Anchor => {
                    let value = arg.text().unwrap_or_default();
                    if matches!(
                        value.to_ascii_lowercase().as_str(),
                        "onecell" | "twocell" | "absolute"
                    ) {
                        geometry.mode = Some(value);
                    } else {
                        geometry.anchor = Some(value);
                    }
                },
                PictureKey::AnchorMode => geometry.mode = arg.text(),
                PictureKey::Title
                | PictureKey::Decorative
                | PictureKey::Hyperlink
                | PictureKey::ContentType
                | PictureKey::FileSize
                | PictureKey::Pending => {},
            }
        }
        let src = src.ok_or_else(|| missing("src", "a picture needs a src"))?;
        let data = read_source(&src)?;
        if !(data.starts_with(b"\x89PNG\r\n\x1a\n") || data.starts_with(&[0xff, 0xd8, 0xff])) {
            return Err(error(
                Code::UnsupportedCapability,
                "the legacy drawing stores PNG and JPEG pictures",
            )
            .with_property("src"));
        }
        let anchor = geometry.anchor(None, [0, 0, 5, 5])?;
        let name = match name {
            Some(name) => name,
            None => format!(
                "Picture {}",
                self.features_of(parent.sheet, Kind::Picture)?.len() + 1
            ),
        };
        let spec = PictureSpec {
            data,
            anchor,
            name: Some(name),
            description,
            rotation,
            flip_h,
            flip_v,
        };
        let (commit, index) = publish!(self, limits, |edit| edit
            .add_picture(parent.sheet, &spec)
            .map_err(native)?);
        Ok((commit, feature_path(parent.sheet, "picture", Some(index))))
    }

    fn add_shape(
        &self,
        parent: &Parent,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<(WriteCommit, Path<'static>)> {
        let mut spec = ShapeSpec {
            geometry: ShapeGeometry::Rectangle,
            anchor: Anchor {
                first_row: 1,
                first_column: 1,
                last_row: 4,
                last_column: 6,
                behavior: litchi_xls::AnchorBehavior::MoveAndResize,
            },
            text: None,
            fill: None,
            line: None,
            name: None,
            description: None,
            rotation: 0.0,
            flip_h: false,
            flip_v: false,
        };
        let mut geometry = Geometry::default();
        for (key, name_key, arg) in admit(&features::SHAPE_TABLE, Operation::Add, props)? {
            let text = || arg.text().unwrap_or_default();
            match key {
                ShapeKey::Name => spec.name = arg.text(),
                ShapeKey::Text => spec.text = Some(text().replace("\\n", "\n")),
                ShapeKey::Geometry => spec.geometry = shape_geometry(name_key, &text())?,
                ShapeKey::X => geometry.x = Some(cells(name_key, &arg)?),
                ShapeKey::Y => geometry.y = Some(cells(name_key, &arg)?),
                ShapeKey::Width => geometry.width = Some(cells(name_key, &arg)?),
                ShapeKey::Height => geometry.height = Some(cells(name_key, &arg)?),
                ShapeKey::Anchor => geometry.anchor = arg.text(),
                ShapeKey::Fill => {
                    let value = text();
                    spec.fill = Some(if value.eq_ignore_ascii_case("none") {
                        None
                    } else {
                        Some(parse_color(name_key, &value, ColorContext::Drawing)?)
                    });
                },
                ShapeKey::Line => {
                    let value = text();
                    spec.line = Some(if value.eq_ignore_ascii_case("none") {
                        None
                    } else {
                        let mut parts = value.split(':');
                        let color = parse_color(
                            name_key,
                            parts.next().unwrap_or(""),
                            ColorContext::Drawing,
                        )?;
                        let width = parts
                            .next()
                            .map(|width| {
                                width
                                    .trim()
                                    .trim_end_matches("pt")
                                    .parse::<f64>()
                                    .ok()
                                    .filter(|points| {
                                        points.is_finite() && *points >= 0.0 && *points <= 1_584.0
                                    })
                                    .map(|points| {
                                        #[allow(
                                            clippy::cast_possible_truncation,
                                            clippy::cast_sign_loss,
                                            reason = "bounded line width"
                                        )]
                                        let emu = (points * 12_700.0).round() as u32;
                                        emu
                                    })
                                    .ok_or_else(|| {
                                        invalid(name_key, "the line width is a number of points")
                                    })
                            })
                            .transpose()?
                            .unwrap_or(9_525);
                        if let Some(dash) = parts.next() {
                            if !dash.trim().eq_ignore_ascii_case("solid") {
                                return Err(error(
                                    Code::UnsupportedProperty,
                                    "the legacy drawing writes solid outlines",
                                )
                                .with_property(name_key));
                            }
                        }
                        Some((color, width))
                    });
                },
                ShapeKey::Rotation => spec.rotation = arg.number(name_key)?,
                ShapeKey::Flip => (spec.flip_h, spec.flip_v) = flip_value(name_key, &text())?,
                ShapeKey::FlipH => spec.flip_h = arg.flag(),
                ShapeKey::FlipV => spec.flip_v = arg.flag(),
                ShapeKey::FlipBoth => (spec.flip_h, spec.flip_v) = (arg.flag(), arg.flag()),
                ShapeKey::Alt => spec.description = arg.text(),
                ShapeKey::Font
                | ShapeKey::Size
                | ShapeKey::Bold
                | ShapeKey::Italic
                | ShapeKey::Underline
                | ShapeKey::Color
                | ShapeKey::Align
                | ShapeKey::Valign
                | ShapeKey::Margin
                | ShapeKey::Hyperlink
                | ShapeKey::Pending => {},
            }
        }
        if !geometry.is_empty() {
            spec.anchor = geometry.anchor(None, [1, 1, 5, 3])?;
        }
        if spec.name.is_none() {
            spec.name = Some(format!(
                "Shape {}",
                self.features_of(parent.sheet, Kind::Shape)?.len() + 1
            ));
        }
        let (commit, index) = publish!(self, limits, |edit| edit
            .add_shape(parent.sheet, &spec)
            .map_err(native)?);
        Ok((commit, feature_path(parent.sheet, "shape", Some(index))))
    }

    fn set_drawing(
        &self,
        sheet: usize,
        feature: &Feature,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<WriteCommit> {
        let Data::Drawing(drawing) = &feature.data else {
            return Err(error(
                Code::InternalError,
                "a drawing node holds no drawing",
            ));
        };
        let mut change = DrawingChange::default();
        let mut geometry = Geometry::default();
        let (mut flip_h, mut flip_v) = (drawing.flip_h, drawing.flip_v);
        let mut flipped = false;
        let picture = feature.kind == Kind::Picture;
        let admitted: Vec<(&'static str, &'static str, Arg)> = if picture {
            admit(&features::PICTURE_TABLE, Operation::Set, props)?
                .into_iter()
                .map(|(key, name, arg)| {
                    let field = match key {
                        PictureKey::Name => "name",
                        PictureKey::Alt => "alt",
                        PictureKey::X => "x",
                        PictureKey::Y => "y",
                        PictureKey::Width => "width",
                        PictureKey::Height => "height",
                        PictureKey::Rotation => "rotation",
                        PictureKey::Flip => "flip",
                        PictureKey::FlipH => "flipH",
                        PictureKey::FlipV => "flipV",
                        PictureKey::FlipBoth => "flipBoth",
                        _ => "",
                    };
                    (field, name, arg)
                })
                .collect()
        } else {
            admit(&features::SHAPE_TABLE, Operation::Set, props)?
                .into_iter()
                .map(|(key, name, arg)| {
                    let field = match key {
                        ShapeKey::Name => "name",
                        ShapeKey::Alt => "alt",
                        ShapeKey::X => "x",
                        ShapeKey::Y => "y",
                        ShapeKey::Width => "width",
                        ShapeKey::Height => "height",
                        ShapeKey::Rotation => "rotation",
                        ShapeKey::Flip => "flip",
                        ShapeKey::FlipH => "flipH",
                        ShapeKey::FlipV => "flipV",
                        ShapeKey::FlipBoth => "flipBoth",
                        _ => "",
                    };
                    (field, name, arg)
                })
                .collect()
        };
        for (field, name, arg) in admitted {
            match field {
                "name" => change.name = Some(arg.text()),
                "alt" => change.description = Some(arg.text()),
                "x" => geometry.x = Some(cells(name, &arg)?),
                "y" => geometry.y = Some(cells(name, &arg)?),
                "width" => geometry.width = Some(cells(name, &arg)?),
                "height" => geometry.height = Some(cells(name, &arg)?),
                "rotation" => change.rotation = Some(arg.number(name)?),
                "flip" => {
                    (flip_h, flip_v) = flip_value(name, &arg.text().unwrap_or_default())?;
                    flipped = true;
                },
                "flipH" => {
                    flip_h = arg.flag();
                    flipped = true;
                },
                "flipV" => {
                    flip_v = arg.flag();
                    flipped = true;
                },
                "flipBoth" => {
                    (flip_h, flip_v) = (arg.flag(), arg.flag());
                    flipped = true;
                },
                _ => {
                    return Err(
                        error(Code::UnsupportedProperty, "the property is not settable")
                            .with_property(name),
                    );
                },
            }
        }
        if flipped {
            change.flip = Some((flip_h, flip_v));
        }
        if !geometry.is_empty() {
            change.anchor = Some(geometry.anchor(drawing.anchor, [0, 0, 5, 5])?);
        }
        let index = feature.index as usize;
        let (commit, ()) = publish!(self, limits, |edit| edit
            .set_drawing(sheet, picture, index, &change)
            .map_err(native)?);
        Ok(commit)
    }
}

// ---------------------------------------------------------------------------
// Conditional formats

/// What a conditional format's rule does, in OfficeCLI's vocabulary.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum CfKind {
    CellIs,
    Formula,
    ColorScale,
    DataBar,
    IconSet,
    TopN,
    AboveAverage,
    DuplicateValues,
    UniqueValues,
    /// A text rule with its `ctp` (0 contains, 1 not contains, 2 begins
    /// with, 3 ends with).
    Text(u16),
    ContainsBlanks,
    NotContainsBlanks,
    ContainsErrors,
    NotContainsErrors,
    TimePeriod,
}

/// A CF element word or `type` value: the rule kind and the flags it implies.
struct CfType {
    kind: CfKind,
    percent: bool,
    bottom: bool,
    below: bool,
}

fn cf_type(value: &str) -> Option<CfType> {
    let plain = |kind| CfType {
        kind,
        percent: false,
        bottom: false,
        below: false,
    };
    Some(match value.trim().to_ascii_lowercase().as_str() {
        "cellis" | "highlight" => plain(CfKind::CellIs),
        "formula" | "expression" | "formulacf" => plain(CfKind::Formula),
        "colorscale" => plain(CfKind::ColorScale),
        "databar" => plain(CfKind::DataBar),
        "iconset" => plain(CfKind::IconSet),
        "top10" | "top" | "topn" => plain(CfKind::TopN),
        "toppercent" => CfType {
            percent: true,
            ..plain(CfKind::TopN)
        },
        "bottom" | "bottomn" => CfType {
            bottom: true,
            ..plain(CfKind::TopN)
        },
        "bottompercent" => CfType {
            percent: true,
            bottom: true,
            ..plain(CfKind::TopN)
        },
        "aboveaverage" => plain(CfKind::AboveAverage),
        "belowaverage" => CfType {
            below: true,
            ..plain(CfKind::AboveAverage)
        },
        "duplicatevalues" => plain(CfKind::DuplicateValues),
        "uniquevalues" => plain(CfKind::UniqueValues),
        "containstext" | "contains" => plain(CfKind::Text(0)),
        "notcontains" | "notcontainstext" => plain(CfKind::Text(1)),
        "beginswith" => plain(CfKind::Text(2)),
        "endswith" => plain(CfKind::Text(3)),
        "containsblanks" => plain(CfKind::ContainsBlanks),
        "notcontainsblanks" => plain(CfKind::NotContainsBlanks),
        "containserrors" => plain(CfKind::ContainsErrors),
        "notcontainserrors" => plain(CfKind::NotContainsErrors),
        "dateoccurring" | "timeperiod" => plain(CfKind::TimePeriod),
        _ => return None,
    })
}

/// The kind of an existing rule.
fn rule_kind(rule: &sf::Rule) -> Option<CfKind> {
    Some(match rule.rule_type {
        RuleType::CellIs => CfKind::CellIs,
        RuleType::ColorScale => CfKind::ColorScale,
        RuleType::DataBar => CfKind::DataBar,
        RuleType::IconSet => CfKind::IconSet,
        RuleType::TopN => CfKind::TopN,
        RuleType::Expression => match rule.template {
            1 => CfKind::Formula,
            7 => CfKind::UniqueValues,
            8 => CfKind::Text(rule.text_operator),
            9 => CfKind::ContainsBlanks,
            10 => CfKind::NotContainsBlanks,
            11 => CfKind::ContainsErrors,
            12 => CfKind::NotContainsErrors,
            15..=24 => CfKind::TimePeriod,
            25 | 26 | 29 | 30 => CfKind::AboveAverage,
            27 => CfKind::DuplicateValues,
            _ => return None,
        },
    })
}

/// The CF element words OfficeCLI adds directly (`--type databar`).
pub(super) fn cf_element(name: &str) -> Option<&'static str> {
    Some(match name.to_ascii_lowercase().as_str() {
        "cf" | "conditionalformatting" => "",
        "databar" => "dataBar",
        "colorscale" => "colorScale",
        "iconset" => "iconSet",
        "formulacf" => "formula",
        "cellis" => "cellIs",
        "topn" => "top10",
        "aboveaverage" => "aboveAverage",
        "duplicatevalues" => "duplicateValues",
        "uniquevalues" => "uniqueValues",
        "containstext" => "containsText",
        "dateoccurring" => "dateOccurring",
        "cfextended" => "",
        _ => return None,
    })
}

/// A cellIs comparison as its `cp` wire value.
fn cf_operator(key: &'static str, value: &str) -> Result<u8> {
    Ok(match value.trim().to_ascii_lowercase().as_str() {
        "between" => 1,
        "notbetween" => 2,
        "equal" | "eq" | "=" | "==" => 3,
        "notequal" | "ne" | "!=" | "<>" => 4,
        "greaterthan" | "gt" | ">" => 5,
        "lessthan" | "lt" | "<" => 6,
        "greaterthanorequal" | "gte" | "ge" | ">=" => 7,
        "lessthanorequal" | "lte" | "le" | "<=" => 8,
        _ => return Err(invalid(key, "the operator is not a cellIs comparison")),
    })
}

/// A time period as its date template.
fn cf_period(key: &'static str, value: &str) -> Result<u16> {
    let value = value.trim();
    features::PERIODS
        .iter()
        .position(|(name, _)| name.eq_ignore_ascii_case(value))
        .and_then(|index| u16::try_from(index).ok())
        .map(|index| index + 15)
        .ok_or_else(|| invalid(key, "the period is not a time period"))
}

/// The top-left cell of the first range, the anchor of generated formulas.
fn first_cell(ranges: &[sf::Area]) -> String {
    ranges.first().map_or_else(
        || "A1".to_owned(),
        |area| {
            crate::format::workbook::Label::cell(
                u32::from(area.first_column),
                u32::from(area.first_row),
            )
            .as_str()
            .to_owned()
        },
    )
}

fn quoted(text: &str) -> String {
    format!("\"{}\"", text.replace('"', "\"\""))
}

/// The formula Excel writes for a text, blank or error rule.
fn generated_formula(kind: CfKind, text: &str, cell: &str) -> Option<String> {
    let text = quoted(text);
    Some(match kind {
        CfKind::Text(0) => format!("NOT(ISERROR(SEARCH({text},{cell})))"),
        CfKind::Text(1) => format!("ISERROR(SEARCH({text},{cell}))"),
        CfKind::Text(2) => format!("LEFT({cell},LEN({text}))={text}"),
        CfKind::Text(_) => format!("RIGHT({cell},LEN({text}))={text}"),
        CfKind::ContainsBlanks => format!("LEN(TRIM({cell}))=0"),
        CfKind::NotContainsBlanks => format!("LEN(TRIM({cell}))>0"),
        CfKind::ContainsErrors => format!("ISERROR({cell})"),
        CfKind::NotContainsErrors => format!("NOT(ISERROR({cell}))"),
        _ => return None,
    })
}

/// The formula Excel writes for a date template.
fn period_formula(template: u16, cell: &str) -> String {
    match template {
        15 => format!("FLOOR({cell},1)=TODAY()"),
        16 => format!("FLOOR({cell},1)=TODAY()+1"),
        17 => format!("FLOOR({cell},1)=TODAY()-1"),
        18 => format!("AND(TODAY()-FLOOR({cell},1)<=6,FLOOR({cell},1)<=TODAY())"),
        // EDATE is an add-in function in BIFF8: the month bounds use DATE.
        19 => format!(
            "AND({cell}>=DATE(YEAR(TODAY()),MONTH(TODAY())-1,1),{cell}<DATE(YEAR(TODAY()),MONTH(TODAY()),1))"
        ),
        20 => format!(
            "AND({cell}>=DATE(YEAR(TODAY()),MONTH(TODAY())+1,1),{cell}<DATE(YEAR(TODAY()),MONTH(TODAY())+2,1))"
        ),
        21 => format!(
            "AND(TODAY()-ROUNDDOWN({cell},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN({cell},0)-TODAY()<=7-WEEKDAY(TODAY()))"
        ),
        22 => format!(
            "AND(ROUNDDOWN({cell},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN({cell},0)-TODAY()<(15-WEEKDAY(TODAY())))"
        ),
        23 => format!(
            "AND(TODAY()-ROUNDDOWN({cell},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN({cell},0)<(WEEKDAY(TODAY())+7))"
        ),
        _ => format!("AND(MONTH({cell})=MONTH(TODAY()),YEAR({cell})=YEAR(TODAY()))"),
    }
}

/// A threshold as `(cfvoType, value)`; `fallback` is the type an absent or
/// `auto` value takes.
fn cf_threshold(
    key: &'static str,
    value: Option<&str>,
    fallback: u8,
) -> Result<(u8, Option<String>)> {
    let value = value
        .map(str::trim)
        .filter(|value| !value.is_empty() && !value.eq_ignore_ascii_case("auto"));
    let Some(value) = value else {
        return Ok((fallback, None));
    };
    let (kind, text) = if let Some(rest) = value.strip_suffix('%') {
        (4, rest)
    } else if let Some((prefix, rest)) = value.split_once(':') {
        let kind = match prefix.to_ascii_lowercase().as_str() {
            "percent" => 4,
            "percentile" => 5,
            "num" | "number" => 1,
            "formula" => 7,
            _ => {
                return Err(invalid(
                    key,
                    "the threshold kind is num, percent, percentile or formula",
                ));
            },
        };
        (kind, rest)
    } else {
        (1, value)
    };
    if kind != 7 && text.trim().parse::<f64>().is_err() {
        return Err(invalid(key, "the threshold is not a number"));
    }
    Ok((kind, Some(text.trim().to_owned())))
}

/// The leaf threshold of `(cfvoType, value)`.
fn cfvo(key: &'static str, (kind, value): &(u8, Option<String>)) -> Result<Cfvo> {
    Ok(match (kind, value) {
        (7, Some(formula)) => Cfvo {
            kind: 7,
            number: None,
            formula: Some(formula_tokens(key, formula)?),
        },
        (2 | 3, _) | (_, None) => Cfvo {
            kind: *kind,
            number: (!matches!(kind, 2 | 3)).then_some(0.0),
            formula: None,
        },
        (kind, Some(number)) => Cfvo {
            kind: *kind,
            number: Some(
                number
                    .parse()
                    .map_err(|_| invalid(key, "the threshold is not a number"))?,
            ),
            formula: None,
        },
    })
}

/// A requested (or current) formatting, as RGB colors.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
struct CfStyle {
    fill: Option<u32>,
    font_color: Option<u32>,
    bold: Option<bool>,
    italic: Option<bool>,
    strike: Option<bool>,
    underline: Option<u8>,
    size_twips: Option<u32>,
    font_name: Option<String>,
}

impl CfStyle {
    fn of(style: &sf::Style) -> Self {
        Self {
            fill: style.fill,
            font_color: style.font_color,
            bold: style.bold,
            italic: style.italic,
            strike: style.strike,
            underline: style.underline,
            size_twips: style.size_twips,
            font_name: style.font_name.clone(),
        }
    }

    fn merge(&mut self, change: &Self) {
        if change.fill.is_some() {
            self.fill = change.fill;
        }
        if change.font_color.is_some() {
            self.font_color = change.font_color;
        }
        if change.bold.is_some() {
            self.bold = change.bold;
        }
        if change.italic.is_some() {
            self.italic = change.italic;
        }
        if change.strike.is_some() {
            self.strike = change.strike;
        }
        if change.underline.is_some() {
            self.underline = change.underline;
        }
        if change.size_twips.is_some() {
            self.size_twips = change.size_twips;
        }
        if change.font_name.is_some() {
            self.font_name.clone_from(&change.font_name);
        }
    }

    fn is_empty(&self) -> bool {
        *self == Self::default()
    }

    /// The leaf style: colors as palette indexes of the workbook.
    fn leaf(&self, palette: &litchi_xls::Palette) -> Result<StyleSpec> {
        let index = |key: &'static str, rgb: u32| -> Result<u8> {
            (8u16..=63)
                .find(|index| {
                    palette.color(*index).is_some_and(|color| {
                        (u32::from(color.red()) << 16)
                            | (u32::from(color.green()) << 8)
                            | u32::from(color.blue())
                            == rgb
                    })
                })
                .and_then(|index| u8::try_from(index).ok())
                .ok_or_else(|| {
                    invalid(key, "the color is not an entry of the workbook palette")
                        .with_suggestion("use one of the workbook's 56 palette colors")
                })
        };
        Ok(StyleSpec {
            fill: self.fill.map(|rgb| index("fill", rgb)).transpose()?,
            font_color: self
                .font_color
                .map(|rgb| index("font.color", rgb).map(u16::from))
                .transpose()?,
            bold: self.bold,
            italic: self.italic,
            strike: self.strike,
            underline: self.underline,
            size_twips: self.size_twips,
            font_name: self.font_name.clone(),
        })
    }
}

/// The parts of a CF request, applied to a fresh or an existing rule.
#[derive(Default)]
struct CfRequest {
    kind: Option<CfType>,
    ranges: Option<Vec<sf::Area>>,
    stop_if_true: Option<bool>,
    operator: Option<u8>,
    value: Option<String>,
    value2: Option<String>,
    formula: Option<String>,
    text: Option<String>,
    rank: Option<u16>,
    percent: Option<bool>,
    bottom: Option<bool>,
    above_average: Option<bool>,
    equal_average: Option<bool>,
    std_dev: Option<u16>,
    period: Option<u16>,
    color: Option<u32>,
    min: Option<String>,
    max: Option<String>,
    min_length: Option<u8>,
    max_length: Option<u8>,
    show_value: Option<bool>,
    min_color: Option<u32>,
    mid_color: Option<u32>,
    max_color: Option<u32>,
    midpoint: Option<String>,
    iconset: Option<u8>,
    reverse: Option<bool>,
    style: CfStyle,
}

fn cf_request(operation: Operation, props: &Props<'_>) -> Result<CfRequest> {
    let mut request = CfRequest::default();
    for (key, name, arg) in admit(&features::CF_TABLE, operation, props)? {
        let text = || arg.text().unwrap_or_default();
        let color = || arg_color(name, &arg, ColorContext::Spreadsheet);
        match key {
            CfKey::Type => {
                request.kind = Some(cf_type(&text()).ok_or_else(|| {
                    invalid(name, "the conditional-format type is not one Excel knows")
                })?);
            },
            CfKey::Ref => request.ranges = Some(parse_ranges(name, &text())?),
            CfKey::StopIfTrue => request.stop_if_true = Some(arg.flag()),
            CfKey::Operator => request.operator = Some(cf_operator(name, &text())?),
            CfKey::Value => request.value = arg.text(),
            CfKey::Value2 => request.value2 = arg.text(),
            CfKey::Formula => {
                let formula = text();
                if formula.starts_with('=') {
                    return Err(invalid(name, "a formula is written without a leading ="));
                }
                request.formula = Some(formula);
            },
            CfKey::Text => request.text = arg.text(),
            CfKey::Rank => {
                request.rank = Some(
                    u16::try_from(arg.int(name)?)
                        .ok()
                        .filter(|rank| (1..=1_000).contains(rank))
                        .ok_or_else(|| invalid(name, "the rank is 1 to 1000"))?,
                );
            },
            CfKey::Percent => request.percent = Some(arg.flag()),
            CfKey::Bottom => request.bottom = Some(arg.flag()),
            CfKey::AboveAverage => request.above_average = Some(arg.flag()),
            CfKey::EqualAverage => request.equal_average = Some(arg.flag()),
            CfKey::StdDev => {
                request.std_dev = Some(
                    u16::try_from(arg.int(name)?)
                        .ok()
                        .filter(|value| (1..=2).contains(value))
                        .ok_or_else(|| invalid(name, "stdDev is 1 or 2 in BIFF8"))?,
                );
            },
            CfKey::Period => request.period = Some(cf_period(name, &text())?),
            CfKey::Color => request.color = Some(color()?),
            CfKey::Min => request.min = arg.text(),
            CfKey::Max => request.max = arg.text(),
            CfKey::MinLength | CfKey::MaxLength => {
                let value = u8::try_from(arg.int(name)?)
                    .ok()
                    .filter(|value| *value <= 100)
                    .ok_or_else(|| invalid(name, "a data-bar length is 0 to 100"))?;
                if key == CfKey::MinLength {
                    request.min_length = Some(value);
                } else {
                    request.max_length = Some(value);
                }
            },
            CfKey::ShowValue => request.show_value = Some(arg.flag()),
            CfKey::MinColor => request.min_color = Some(color()?),
            CfKey::MidColor => request.mid_color = Some(color()?),
            CfKey::MaxColor => request.max_color = Some(color()?),
            CfKey::Midpoint => request.midpoint = arg.text(),
            CfKey::Iconset => {
                let value = text();
                let value = value.trim();
                request.iconset = Some(
                    features::ICON_SETS
                        .iter()
                        .position(|set| set.eq_ignore_ascii_case(value))
                        .and_then(|index| u8::try_from(index).ok())
                        .ok_or_else(|| invalid(name, "the icon set is not a BIFF8 icon set"))?,
                );
            },
            CfKey::Reverse => request.reverse = Some(arg.flag()),
            CfKey::Fill => request.style.fill = Some(color()?),
            CfKey::FontColor => request.style.font_color = Some(color()?),
            CfKey::FontBold => request.style.bold = Some(arg.flag()),
            CfKey::FontItalic => request.style.italic = Some(arg.flag()),
            CfKey::FontStrike => request.style.strike = Some(arg.flag()),
            CfKey::FontUnderline => {
                request.style.underline = match text().to_ascii_lowercase().as_str() {
                    "single" | "true" => Some(1),
                    "double" => Some(2),
                    "none" | "false" => Some(0),
                    _ => return Err(invalid(name, "underline is single, double or none")),
                };
            },
            CfKey::FontSize => {
                if let Arg::Length(length) = arg {
                    let twips = (length.points() * 20.0).round();
                    if !(20.0..=8_191.0).contains(&twips) {
                        return Err(invalid(name, "the font size is 1 to 409.55 points"));
                    }
                    #[allow(
                        clippy::cast_possible_truncation,
                        clippy::cast_sign_loss,
                        reason = "bounded above"
                    )]
                    let twips = twips as u32;
                    request.style.size_twips = Some(twips);
                }
            },
            CfKey::FontName => request.style.font_name = arg.text(),
            CfKey::Priority
            | CfKey::DxfId
            | CfKey::NegativeColor
            | CfKey::AxisColor
            | CfKey::AxisPosition
            | CfKey::Direction => {},
        }
    }
    Ok(request)
}

/// The state of one rule as the adapter edits it: a fresh rule of a kind,
/// or the projection of an existing first rule.
#[derive(Debug, Clone)]
struct CfState {
    kind: CfKind,
    stop_if_true: bool,
    operator: u8,
    formulas: Vec<String>,
    text: Option<String>,
    rank: u16,
    percent: bool,
    bottom: bool,
    below: bool,
    equal: bool,
    std_dev: u16,
    template: u16,
    bar: Option<(
        (u8, Option<String>),
        (u8, Option<String>),
        u32,
        u8,
        u8,
        bool,
    )>,
    scale: Option<Vec<((u8, Option<String>), u32)>>,
    icons: Option<(u8, Vec<(u8, Option<String>)>, bool, bool)>,
    style: CfStyle,
}

impl CfState {
    fn fresh(kind: CfKind) -> Self {
        Self {
            kind,
            stop_if_true: false,
            operator: 5,
            formulas: Vec::new(),
            text: None,
            rank: 10,
            percent: false,
            bottom: false,
            below: false,
            equal: false,
            std_dev: 0,
            template: if kind == CfKind::TimePeriod { 15 } else { 0 },
            bar: None,
            scale: None,
            icons: None,
            style: CfStyle::default(),
        }
    }

    fn of(rule: &sf::Rule) -> Option<Self> {
        let kind = rule_kind(rule)?;
        let threshold = |value: &sf::Threshold| (value.kind, value.value.clone());
        Some(Self {
            kind,
            stop_if_true: rule.stop_if_true,
            operator: rule.operator.unwrap_or(5),
            formulas: rule.formulas.clone(),
            text: rule.text.clone(),
            rank: if rule.rule_type == RuleType::TopN {
                rule.parameter
            } else {
                10
            },
            percent: rule.percent,
            bottom: rule.bottom,
            below: matches!(rule.template, 26 | 30),
            equal: matches!(rule.template, 29 | 30),
            std_dev: if matches!(rule.template, 25 | 26 | 29 | 30) {
                rule.parameter
            } else {
                0
            },
            template: rule.template,
            bar: rule.data_bar.as_ref().map(|bar| {
                (
                    threshold(&bar.min),
                    threshold(&bar.max),
                    bar.color.unwrap_or(0x0063_8EC6),
                    bar.min_length,
                    bar.max_length,
                    bar.show_value,
                )
            }),
            scale: rule.color_scale.as_ref().map(|scale| {
                scale
                    .points
                    .iter()
                    .map(|(point, color)| (threshold(point), color.unwrap_or(0)))
                    .collect()
            }),
            icons: rule.icon_set.as_ref().map(|icons| {
                (
                    icons.icon_set,
                    icons.thresholds.iter().map(threshold).collect(),
                    icons.reverse,
                    icons.show_value,
                )
            }),
            style: CfStyle::of(&rule.style),
        })
    }

    /// Apply a request (`fresh` for an add).
    fn apply(&mut self, request: &CfRequest, ranges: &[sf::Area], fresh: bool) -> Result<()> {
        if let Some(value) = request.stop_if_true {
            self.stop_if_true = value;
        }
        let moved = request.ranges.is_some();
        match self.kind {
            CfKind::CellIs => {
                if let Some(operator) = request.operator {
                    self.operator = operator;
                }
                if let Some(value) = &request.value {
                    if self.formulas.is_empty() {
                        self.formulas.push(value.clone());
                    } else {
                        self.formulas[0].clone_from(value);
                    }
                }
                if let Some(value) = &request.value2 {
                    self.formulas.truncate(1);
                    self.formulas.push(value.clone());
                }
                let between = matches!(self.operator, 1 | 2);
                if !between {
                    self.formulas.truncate(1);
                }
                if self.formulas.is_empty() {
                    return Err(missing("value", "a cellIs rule needs a value"));
                }
                if between && self.formulas.len() < 2 {
                    return Err(missing("value2", "a between rule needs value2"));
                }
            },
            CfKind::Formula => {
                if let Some(formula) = request.formula.clone().or_else(|| request.value.clone()) {
                    self.formulas = vec![formula];
                }
                if self.formulas.is_empty() {
                    return Err(missing("formula", "a formula rule needs a formula"));
                }
            },
            CfKind::TopN => {
                if let Some(rank) = request.rank.or_else(|| {
                    request
                        .value
                        .as_deref()
                        .and_then(|value| value.parse().ok())
                }) {
                    self.rank = rank;
                }
                if let Some(value) = request.percent {
                    self.percent = value;
                }
                if let Some(value) = request.bottom {
                    self.bottom = value;
                }
            },
            CfKind::AboveAverage => {
                if let Some(above) = request.above_average {
                    self.below = !above;
                }
                if let Some(value) = request.std_dev {
                    self.std_dev = value;
                }
                if let Some(value) = request.equal_average {
                    self.equal = value;
                }
                // There is no equal-average template with deviations.
                if self.std_dev > 0 {
                    self.equal = false;
                }
            },
            CfKind::Text(_) => {
                if let Some(text) = &request.text {
                    self.text = Some(text.clone());
                }
                let text = self
                    .text
                    .clone()
                    .ok_or_else(|| missing("text", "a text rule needs text"))?;
                if fresh || request.text.is_some() || moved {
                    self.formulas = vec![
                        generated_formula(self.kind, &text, &first_cell(ranges))
                            .ok_or_else(|| error(Code::InternalError, "no generated formula"))?,
                    ];
                }
            },
            CfKind::ContainsBlanks
            | CfKind::NotContainsBlanks
            | CfKind::ContainsErrors
            | CfKind::NotContainsErrors => {
                if fresh || moved {
                    self.formulas = vec![
                        generated_formula(self.kind, "", &first_cell(ranges))
                            .ok_or_else(|| error(Code::InternalError, "no generated formula"))?,
                    ];
                }
            },
            CfKind::TimePeriod => {
                if let Some(template) = request.period {
                    self.template = template;
                }
                if fresh || request.period.is_some() || moved {
                    self.formulas = vec![period_formula(self.template, &first_cell(ranges))];
                }
            },
            CfKind::DataBar => {
                let mut bar = match self.bar.take() {
                    Some(bar) => bar,
                    None => ((2, None), (3, None), 0x0063_8EC6, 10, 90, true),
                };
                if request.min.is_some() {
                    bar.0 = cf_threshold("min", request.min.as_deref(), 2)?;
                }
                if request.max.is_some() {
                    bar.1 = cf_threshold("max", request.max.as_deref(), 3)?;
                }
                if let Some(color) = request.color {
                    bar.2 = color;
                }
                if let Some(value) = request.min_length {
                    bar.3 = value;
                }
                if let Some(value) = request.max_length {
                    bar.4 = value;
                }
                if let Some(value) = request.show_value {
                    bar.5 = value;
                }
                self.bar = Some(bar);
            },
            CfKind::ColorScale => {
                let mut scale = match self.scale.take() {
                    Some(scale) => scale,
                    None => vec![((2, None), 0x00F8_696B), ((3, None), 0x0063_BE7B)],
                };
                if (request.mid_color.is_some() || request.midpoint.is_some()) && scale.len() == 2 {
                    scale.insert(1, ((5, Some("50".to_owned())), 0x00FF_EB84));
                }
                if let Some(midpoint) = &request.midpoint {
                    let mut threshold = cf_threshold("midpoint", Some(midpoint), 5)?;
                    if threshold.0 == 1 && !midpoint.contains(':') {
                        threshold.0 = 5;
                    }
                    scale[1].0 = threshold;
                }
                if let Some(color) = request.mid_color {
                    scale[1].1 = color;
                }
                if let Some(color) = request.min_color {
                    scale[0].1 = color;
                }
                if let Some(color) = request.max_color {
                    let last = scale.len() - 1;
                    scale[last].1 = color;
                }
                self.scale = Some(scale);
            },
            CfKind::IconSet => {
                let mut icons = match self.icons.take() {
                    Some(icons) => icons,
                    None => (0, Vec::new(), false, true),
                };
                if let Some(set) = request.iconset {
                    icons.0 = set;
                    icons.1.clear();
                }
                if icons.1.is_empty() {
                    let count: u16 = match icons.0 {
                        0..=7 => 3,
                        8..=12 => 4,
                        _ => 5,
                    };
                    icons.1 = (0..count)
                        .map(|step| (4, Some((step * 100 / count).to_string())))
                        .collect();
                }
                if let Some(value) = request.reverse {
                    icons.2 = value;
                }
                if let Some(value) = request.show_value {
                    icons.3 = value;
                }
                self.icons = Some(icons);
            },
            CfKind::DuplicateValues | CfKind::UniqueValues => {},
        }
        Ok(())
    }

    /// Whether the rule carries formatting (visual rules do not).
    const fn takes_style(&self) -> bool {
        !matches!(
            self.kind,
            CfKind::ColorScale | CfKind::DataBar | CfKind::IconSet
        )
    }

    /// The leaf format.
    fn format(&self, ranges: Vec<sf::Area>, palette: &litchi_xls::Palette) -> Result<Format> {
        let first = |key: &'static str, formula: Option<&String>| -> Result<Vec<u8>> {
            formula_tokens(key, formula.map_or("", String::as_str))
        };
        let template_formula = |template: u16, formula: Option<&String>| -> Result<FormatRule> {
            Ok(FormatRule::Template {
                template,
                params: [0; 16],
                formula: first("formula", formula)?,
            })
        };
        let rule = match self.kind {
            CfKind::CellIs => FormatRule::CellIs {
                operator: self.operator,
                formulas: [
                    first("value", self.formulas.first())?,
                    match self.formulas.get(1) {
                        Some(formula) => formula_tokens("value2", formula)?,
                        None => Vec::new(),
                    },
                ],
            },
            CfKind::Formula => FormatRule::Expression(first("formula", self.formulas.first())?),
            CfKind::TopN => FormatRule::TopN {
                top: !self.bottom,
                percent: self.percent,
                rank: self.rank,
            },
            CfKind::AboveAverage => {
                let template = match (self.below, self.equal) {
                    (false, false) => 0x19,
                    (true, false) => 0x1a,
                    (false, true) => 0x1d,
                    (true, true) => 0x1e,
                };
                let mut params = [0u8; 16];
                params[0..2].copy_from_slice(&self.std_dev.to_le_bytes());
                // Excel 2007 evaluates an average rule by its template; the
                // formula is a placeholder older readers never see (they
                // skip CF12).
                FormatRule::Template {
                    template,
                    params,
                    formula: formula_tokens("aboveAverage", "TRUE")?,
                }
            },
            CfKind::DuplicateValues => template_formula(0x1b, Some(&"TRUE".to_owned()))?,
            CfKind::UniqueValues => template_formula(7, Some(&"TRUE".to_owned()))?,
            CfKind::Text(operator) => {
                let mut params = [0u8; 16];
                params[0..2].copy_from_slice(&operator.to_le_bytes());
                FormatRule::Template {
                    template: 8,
                    params,
                    formula: first("text", self.formulas.first())?,
                }
            },
            CfKind::ContainsBlanks => template_formula(9, self.formulas.first())?,
            CfKind::NotContainsBlanks => template_formula(10, self.formulas.first())?,
            CfKind::ContainsErrors => template_formula(11, self.formulas.first())?,
            CfKind::NotContainsErrors => template_formula(12, self.formulas.first())?,
            CfKind::TimePeriod => {
                let mut params = [0u8; 16];
                let operator = self
                    .template
                    .checked_sub(15)
                    .and_then(|index| features::PERIODS.get(usize::from(index)))
                    .map_or(0, |(_, operator)| *operator);
                params[0..2].copy_from_slice(&u16::try_from(operator).unwrap_or(0).to_le_bytes());
                FormatRule::Template {
                    template: self.template,
                    params,
                    formula: first("period", self.formulas.first())?,
                }
            },
            CfKind::DataBar => {
                let bar = self
                    .bar
                    .as_ref()
                    .ok_or_else(|| error(Code::InternalError, "no data bar"))?;
                if bar.3 >= bar.4 {
                    return Err(invalid(
                        "maxLength",
                        "a data bar's maxLength exceeds its minLength",
                    ));
                }
                FormatRule::DataBar {
                    color: bar.2,
                    min: cfvo("min", &bar.0)?,
                    max: cfvo("max", &bar.1)?,
                    min_length: bar.3,
                    max_length: bar.4,
                    show_value: bar.5,
                }
            },
            CfKind::ColorScale => FormatRule::ColorScale(
                self.scale
                    .as_ref()
                    .ok_or_else(|| error(Code::InternalError, "no color scale"))?
                    .iter()
                    .map(|(threshold, color)| Ok((cfvo("midpoint", threshold)?, *color)))
                    .collect::<Result<_>>()?,
            ),
            CfKind::IconSet => {
                let icons = self
                    .icons
                    .as_ref()
                    .ok_or_else(|| error(Code::InternalError, "no icon set"))?;
                FormatRule::IconSet {
                    set: icons.0,
                    thresholds: icons
                        .1
                        .iter()
                        .map(|threshold| cfvo("iconset", threshold))
                        .collect::<Result<_>>()?,
                    reverse: icons.2,
                    show_value: icons.3,
                }
            },
        };
        let style = if self.takes_style() {
            self.style.leaf(palette)?
        } else {
            StyleSpec::default()
        };
        Ok(Format {
            ranges,
            rule,
            style,
            stop_if_true: self.stop_if_true && self.takes_style(),
        })
    }
}

impl Xls {
    fn add_cf(
        &self,
        parent: &Parent,
        implied: Option<&'static str>,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<(WriteCommit, Path<'static>)> {
        let mut request = cf_request(Operation::Add, props)?;
        let kind = match (
            request.kind.take(),
            implied.filter(|implied| !implied.is_empty()),
        ) {
            (Some(kind), _) => kind,
            (None, Some(implied)) => cf_type(implied)
                .ok_or_else(|| error(Code::InternalError, "unknown implied type"))?,
            // OfficeCLI's `--type cf` without a type is a data bar.
            (None, None) => {
                cf_type("dataBar").ok_or_else(|| error(Code::InternalError, "no data bar"))?
            },
        };
        let ranges = match request.ranges.take() {
            Some(ranges) => ranges,
            None => parse_ranges("ref", parent.reference.as_deref().unwrap_or("A1:A10"))?,
        };
        let mut state = CfState::fresh(kind.kind);
        state.percent = kind.percent;
        state.bottom = kind.bottom;
        state.below = kind.below;
        state.apply(&request, &ranges, true)?;
        if !request.style.is_empty() && !state.takes_style() {
            return Err(invalid(
                "fill",
                "a color scale, data bar or icon set takes no formatting",
            ));
        }
        state.style = request.style.clone();
        let palette = self.owner().palette()?;
        let format = state.format(ranges, &palette)?;
        let (commit, index) = publish!(self, limits, |edit| edit
            .add_format(parent.sheet, &format)
            .map_err(native)?);
        Ok((commit, feature_path(parent.sheet, "cf", Some(index))))
    }

    fn set_cf(
        &self,
        sheet: usize,
        feature: &Feature,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<WriteCommit> {
        let Data::Cf(cf) = &feature.data else {
            return Err(error(
                Code::InternalError,
                "a conditional-format node holds no format",
            ));
        };
        let request = cf_request(Operation::Set, props)?;
        let index = feature.index as usize;
        // Ranges alone: the rules keep their bytes.
        let only_ranges = props.iter().all(|(key, _)| {
            ["ref", "sqref", "range"]
                .iter()
                .any(|name| key.as_str().eq_ignore_ascii_case(name))
        });
        if only_ranges {
            let ranges = request.ranges.clone().unwrap_or_else(|| cf.ranges.clone());
            let (commit, ()) = publish!(self, limits, |edit| edit
                .set_format_ranges(sheet, index, &ranges)
                .map_err(native)?);
            return Ok(commit);
        }
        let rule = cf.rule.as_ref().ok_or_else(|| {
            error(
                Code::UnsupportedCapability,
                "the conditional format has no rule",
            )
        })?;
        let mut state = CfState::of(rule).ok_or_else(|| {
            error(
                Code::UnsupportedCapability,
                "the conditional-format rule template is not one the XLS editor authors",
            )
        })?;
        let ranges = request.ranges.clone().unwrap_or_else(|| cf.ranges.clone());
        state.apply(&request, &ranges, false)?;
        if !request.style.is_empty() {
            if !state.takes_style() {
                return Err(invalid(
                    "fill",
                    "a color scale, data bar or icon set takes no formatting",
                ));
            }
            state.style.merge(&request.style);
        }
        let palette = self.owner().palette()?;
        let format = state.format(ranges, &palette)?;
        let (commit, ()) = publish!(self, limits, |edit| edit
            .set_format(sheet, index, &format)
            .map_err(native)?);
        Ok(commit)
    }
}

// ---------------------------------------------------------------------------
// Data validations

/// A validation formula's tokens: a list's literal options (`a,b`) become a
/// string list, anything else (a leading `=` dropped) a formula.
fn validation_formula(key: &'static str, value: &str, list: bool) -> Result<(Vec<u8>, bool)> {
    let value = value.trim();
    if value.encode_utf16().count() > 255 {
        return Err(invalid(
            key,
            "a validation formula holds at most 255 characters",
        ));
    }
    if let Some(rest) = value.strip_prefix('=') {
        return Ok((formula_tokens(key, rest)?, false));
    }
    if list {
        let literal = value
            .strip_prefix('"')
            .and_then(|rest| rest.strip_suffix('"'))
            .unwrap_or(value);
        if literal.contains('"') {
            return Err(invalid(key, "list options cannot contain quotes"));
        }
        let looks_like_reference = parse_area(literal).is_some() && !literal.contains(',');
        if looks_like_reference && !value.starts_with('"') {
            return Ok((formula_tokens(key, literal)?, false));
        }
        let units: Vec<u16> = literal.replace(',', "\0").encode_utf16().collect();
        let count = u8::try_from(units.len())
            .map_err(|_| invalid(key, "a list holds at most 255 characters"))?;
        let mut tokens = vec![0x17, count];
        if units.iter().all(|unit| *unit <= 0xff) {
            tokens.push(0);
            #[allow(clippy::cast_possible_truncation, reason = "checked above")]
            tokens.extend(units.iter().map(|unit| *unit as u8));
        } else {
            tokens.push(1);
            for unit in units {
                tokens.extend_from_slice(&unit.to_le_bytes());
            }
        }
        return Ok((tokens, true));
    }
    Ok((formula_tokens(key, value)?, false))
}

/// The state of a validation as the adapter edits it.
#[derive(Debug, Clone)]
struct DvState {
    ranges: Vec<sf::Area>,
    kind: u8,
    operator: u8,
    error_style: u8,
    formula1: Option<String>,
    formula2: Option<String>,
    allow_blank: bool,
    dropdown: bool,
    show_input: bool,
    show_error: bool,
    prompt_title: Option<String>,
    prompt: Option<String>,
    error_title: Option<String>,
    error: Option<String>,
}

impl DvState {
    fn of(validation: &sf::Validation) -> Self {
        Self {
            ranges: validation.ranges.clone(),
            kind: validation.kind,
            operator: validation.operator,
            error_style: validation.error_style,
            formula1: validation.formula1.clone(),
            formula2: validation.formula2.clone(),
            allow_blank: validation.allow_blank,
            dropdown: !validation.suppress_dropdown,
            show_input: validation.show_input,
            show_error: validation.show_error,
            prompt_title: validation.prompt_title.clone(),
            prompt: validation.prompt.clone(),
            error_title: validation.error_title.clone(),
            error: validation.error.clone(),
        }
    }

    fn apply(&mut self, props: &Props<'_>, operation: Operation) -> Result<()> {
        for (key, name, arg) in admit(&features::VALIDATION_TABLE, operation, props)? {
            let text = || arg.text().unwrap_or_default();
            match key {
                ValidationKey::Type => {
                    let value = text();
                    let value = match value.to_ascii_lowercase().as_str() {
                        "any" => "none".to_owned(),
                        other => other.to_owned(),
                    };
                    self.kind = features::VALIDATION_TYPES
                        .iter()
                        .position(|known| known.eq_ignore_ascii_case(&value))
                        .and_then(|index| u8::try_from(index).ok())
                        .ok_or_else(|| {
                            invalid(name, "the validation type is not one Excel knows")
                        })?;
                },
                ValidationKey::Ref => self.ranges = parse_ranges(name, &text())?,
                ValidationKey::Operator => {
                    let value = text();
                    self.operator = features::VALIDATION_OPERATORS
                        .iter()
                        .position(|known| known.eq_ignore_ascii_case(value.trim()))
                        .and_then(|index| u8::try_from(index).ok())
                        .ok_or_else(|| {
                            invalid(name, "the operator is not a validation comparison")
                        })?;
                },
                ValidationKey::Formula1 => self.formula1 = arg.text(),
                ValidationKey::Formula2 => self.formula2 = arg.text(),
                ValidationKey::AllowBlank => self.allow_blank = arg.flag(),
                ValidationKey::ShowError => self.show_error = arg.flag(),
                ValidationKey::ShowInput => self.show_input = arg.flag(),
                ValidationKey::ErrorTitle => self.error_title = arg.text(),
                ValidationKey::Error => self.error = arg.text(),
                ValidationKey::PromptTitle => self.prompt_title = arg.text(),
                ValidationKey::Prompt => self.prompt = arg.text(),
                ValidationKey::ErrorStyle => {
                    let value = text().to_ascii_lowercase();
                    let value = match value.as_str() {
                        "warn" => "warning",
                        "info" => "information",
                        other => other,
                    };
                    self.error_style = features::ERROR_STYLES
                        .iter()
                        .position(|known| known.eq_ignore_ascii_case(value))
                        .and_then(|index| u8::try_from(index).ok())
                        .ok_or_else(|| {
                            invalid(name, "errorStyle is stop, warning or information")
                        })?;
                },
                ValidationKey::InCellDropdown => self.dropdown = arg.flag(),
            }
        }
        for (formula, key) in [(&self.formula1, "formula1"), (&self.formula2, "formula2")] {
            if formula
                .as_deref()
                .is_some_and(|formula| formula.trim().is_empty())
            {
                return Err(invalid(key, "a validation formula cannot be empty"));
            }
        }
        for (text, key, limit) in [
            (&self.prompt_title, "promptTitle", 32),
            (&self.error_title, "errorTitle", 32),
            (&self.prompt, "prompt", 255),
            (&self.error, "error", 225),
        ] {
            if text
                .as_ref()
                .is_some_and(|text| text.encode_utf16().count() > limit)
            {
                return Err(invalid(key, "the validation message is too long for BIFF8"));
            }
        }
        Ok(())
    }

    fn leaf(&self) -> Result<ValidationSpec> {
        let list = self.kind == 3;
        let (formula1, string_list) = match &self.formula1 {
            Some(formula) => validation_formula("formula1", formula, list)?,
            None if self.kind == 0 => (Vec::new(), false),
            None => return Err(missing("formula1", "a typed validation needs formula1")),
        };
        let formula2 = match &self.formula2 {
            Some(formula) if !matches!(self.kind, 3 | 7) => {
                validation_formula("formula2", formula, false)?.0
            },
            _ => Vec::new(),
        };
        Ok(ValidationSpec {
            ranges: self.ranges.clone(),
            kind: self.kind,
            operator: self.operator,
            error_style: self.error_style,
            formula1,
            formula2,
            list: string_list,
            allow_blank: self.allow_blank,
            suppress_dropdown: !self.dropdown,
            show_input: self.show_input,
            show_error: self.show_error,
            prompt_title: self.prompt_title.clone(),
            prompt: self.prompt.clone(),
            error_title: self.error_title.clone(),
            error: self.error.clone(),
        })
    }
}

fn overlaps(left: &[sf::Area], right: &[sf::Area]) -> bool {
    left.iter().any(|a| right.iter().any(|b| a.overlaps(*b)))
}

impl Xls {
    fn add_validation(
        &self,
        parent: &Parent,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<(WriteCommit, Path<'static>)> {
        let reference = props.iter().any(|(key, _)| {
            ["ref", "sqref", "range"]
                .iter()
                .any(|name| key.as_str().eq_ignore_ascii_case(name))
        });
        let fallback = match (&parent.reference, reference) {
            (_, true) => "A1".to_owned(),
            (Some(reference), false) => reference.clone(),
            (None, false) => return Err(missing("ref", "a data validation needs a ref")),
        };
        let mut state = DvState {
            ranges: parse_ranges("ref", &fallback)?,
            kind: 0,
            operator: 0,
            error_style: 0,
            formula1: None,
            formula2: None,
            allow_blank: true,
            dropdown: true,
            show_input: true,
            show_error: true,
            prompt_title: None,
            prompt: None,
            error_title: None,
            error: None,
        };
        state.apply(props, Operation::Add)?;
        let existing = self.features_of(parent.sheet, Kind::DataValidation)?;
        if existing.iter().any(|feature| match &feature.data {
            Data::Validation(held) => overlaps(&held.ranges, &state.ranges),
            _ => false,
        }) {
            return Err(error(
                Code::InvalidArgument,
                "the range overlaps an existing data validation",
            )
            .with_property("ref"));
        }
        let spec = state.leaf()?;
        let (commit, index) = publish!(self, limits, |edit| edit
            .add_validation(parent.sheet, &spec)
            .map_err(native)?);
        Ok((
            commit,
            feature_path(parent.sheet, "dataValidation", Some(index)),
        ))
    }

    fn set_validation(
        &self,
        sheet: usize,
        feature: &Feature,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<WriteCommit> {
        let Data::Validation(current) = &feature.data else {
            return Err(error(
                Code::InternalError,
                "a validation node holds no validation",
            ));
        };
        let mut state = DvState::of(current);
        state.apply(props, Operation::Set)?;
        let index = feature.index as usize;
        let existing = self.features_of(sheet, Kind::DataValidation)?;
        if existing.iter().any(|other| {
            other.index as usize != index
                && match &other.data {
                    Data::Validation(held) => overlaps(&held.ranges, &state.ranges),
                    _ => false,
                }
        }) {
            return Err(error(
                Code::InvalidArgument,
                "the range overlaps another data validation",
            )
            .with_property("ref"));
        }
        let spec = state.leaf()?;
        let (commit, ()) = publish!(self, limits, |edit| edit
            .set_validation(sheet, index, &spec)
            .map_err(native)?);
        Ok(commit)
    }
}

// ---------------------------------------------------------------------------
// The auto-filter

fn filter_range(key: &'static str, value: &str) -> Result<sf::Area> {
    parse_area(value).ok_or_else(|| invalid(key, "the range is not an A1 range"))
}

impl Xls {
    fn add_autofilter(
        &self,
        parent: &Parent,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<(WriteCommit, Path<'static>)> {
        let mut range = parent.reference.clone();
        for (key, _, arg) in admit(&features::AUTOFILTER_TABLE, Operation::Add, props)? {
            if key == AutofilterKey::Range {
                range = arg.text();
            }
        }
        let range = range.ok_or_else(|| missing("range", "an auto-filter needs a range"))?;
        let area = filter_range("range", &range)?;
        let (commit, ()) = publish!(self, limits, |edit| edit
            .set_auto_filter(parent.sheet, Some(area))
            .map_err(native)?);
        Ok((commit, feature_path(parent.sheet, "autofilter", None)))
    }

    fn set_autofilter(
        &self,
        sheet: usize,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<WriteCommit> {
        let current = self
            .features_of(sheet, Kind::Autofilter)?
            .into_iter()
            .find_map(|feature| match feature.data {
                Data::Autofilter(filter) => Some(filter),
                _ => None,
            })
            .ok_or_else(|| error(Code::PathNotFound, "the worksheet has no auto-filter"))?;
        let mut area = current.area;
        for (key, name, arg) in admit(&features::AUTOFILTER_TABLE, Operation::Set, props)? {
            if key == AutofilterKey::Range {
                area = Some(filter_range(name, &arg.text().unwrap_or_default())?);
            }
        }
        let area = area.ok_or_else(|| missing("range", "the auto-filter has no range to keep"))?;
        let (commit, ()) = publish!(self, limits, |edit| edit
            .set_auto_filter(sheet, Some(area))
            .map_err(native)?);
        Ok(commit)
    }
}

// ---------------------------------------------------------------------------
// Page breaks

fn break_index(kind: Kind, key: &'static str, arg: &Arg) -> Result<u16> {
    let (maximum, letters) = if kind == Kind::RowBreak {
        (65_535_i64, false)
    } else {
        (255, true)
    };
    let value = match arg {
        Arg::Text(text)
            if letters && text.trim().bytes().all(|byte| byte.is_ascii_alphabetic()) =>
        {
            i64::from(
                litchi_core::selector::CellRef::parse(&format!("{}1", text.trim()))
                    .ok_or_else(|| invalid(key, "the column is out of range"))?
                    .column(),
            ) + 1
        },
        arg => arg.int(key)?,
    };
    u16::try_from(value)
        .ok()
        .filter(|value| (1..=maximum).contains(&i64::from(*value)))
        .ok_or_else(|| invalid(key, "the break is out of the worksheet"))
}

impl Xls {
    fn breaks(&self, sheet: usize, kind: Kind) -> Result<Vec<PageBreak>> {
        Ok(self
            .features_of(sheet, kind)?
            .into_iter()
            .filter_map(|feature| match feature.data {
                Data::Break(item) => Some(item),
                _ => None,
            })
            .collect())
    }

    fn publish_breaks(
        &self,
        sheet: usize,
        kind: Kind,
        breaks: &[PageBreak],
        limits: &SemanticLimits,
    ) -> Result<WriteCommit> {
        let (commit, ()) = publish!(self, limits, |edit| if kind == Kind::RowBreak {
            edit.set_row_breaks(sheet, breaks).map_err(native)?;
        } else {
            edit.set_column_breaks(sheet, breaks).map_err(native)?;
        });
        Ok(commit)
    }

    fn break_spec(
        kind: Kind,
        current: Option<PageBreak>,
        props: &Props<'_>,
        operation: Operation,
    ) -> Result<PageBreak> {
        let full: u16 = if kind == Kind::RowBreak { 255 } else { 65_535 };
        let mut index = current.map(|item| item.index);
        let mut minimum = current.map_or(0, |item| item.min);
        let mut maximum = current.map_or(full, |item| item.max);
        let admitted = if kind == Kind::RowBreak {
            admit(&features::ROW_BREAK_TABLE, operation, props)?
        } else {
            admit(&features::COL_BREAK_TABLE, operation, props)?
        };
        for (key, name, arg) in admitted {
            match key {
                BreakKey::Index => index = Some(break_index(kind, name, &arg)?),
                BreakKey::Manual => {
                    if !arg.flag() {
                        return Err(invalid(
                            "manual",
                            "an XLS worksheet stores manual page breaks only",
                        ));
                    }
                },
                BreakKey::Min | BreakKey::Max => {
                    let value = u16::try_from(arg.int(name)?)
                        .map_err(|_| invalid(name, "the bound is out of range"))?;
                    if key == BreakKey::Min {
                        minimum = value;
                    } else {
                        maximum = value;
                    }
                },
            }
        }
        let key = if kind == Kind::RowBreak { "row" } else { "col" };
        let index = index.ok_or_else(|| missing(key, "a page break needs its row or column"))?;
        if minimum > maximum || maximum > full {
            return Err(invalid("max", "the break span is outside the worksheet"));
        }
        Ok(PageBreak {
            index,
            min: minimum,
            max: maximum,
        })
    }

    fn add_break(
        &self,
        parent: &Parent,
        kind: Kind,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<(WriteCommit, Path<'static>)> {
        let item = Self::break_spec(kind, None, props, Operation::Add)?;
        let mut breaks = self.breaks(parent.sheet, kind)?;
        if breaks.iter().any(|held| held.index == item.index) {
            return Err(error(
                Code::InvalidArgument,
                "the worksheet already has this page break",
            ));
        }
        let index = breaks.partition_point(|held| held.index < item.index);
        breaks.insert(index, item);
        let commit = self.publish_breaks(parent.sheet, kind, &breaks, limits)?;
        Ok((
            commit,
            feature_path(parent.sheet, features::step_name(kind), Some(index)),
        ))
    }

    fn set_break(
        &self,
        sheet: usize,
        feature: &Feature,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<WriteCommit> {
        let Data::Break(current) = feature.data else {
            return Err(error(
                Code::InternalError,
                "a page-break node holds no break",
            ));
        };
        let item = Self::break_spec(feature.kind, Some(current), props, Operation::Set)?;
        let mut breaks = self.breaks(sheet, feature.kind)?;
        let index = feature.index as usize;
        if index >= breaks.len() {
            return Err(error(Code::PathNotFound, "the page break is missing"));
        }
        breaks.remove(index);
        if breaks.iter().any(|held| held.index == item.index) {
            return Err(error(
                Code::InvalidArgument,
                "the worksheet already has this page break",
            ));
        }
        let at = breaks.partition_point(|held| held.index < item.index);
        breaks.insert(at, item);
        self.publish_breaks(sheet, feature.kind, &breaks, limits)
    }
}

/// The worksheet feature an element word names (`comment`, `note`,
/// `validation`, `image`, `textbox`, a CF word, `pagebreak`), with the rule
/// type a CF word implies.
pub(super) fn feature_kind(name: &str, props: &Props<'_>) -> Option<(Kind, Option<&'static str>)> {
    if let Some(implied) = cf_element(name) {
        return Some((Kind::Cf, Some(implied)));
    }
    if name.eq_ignore_ascii_case("pagebreak") {
        // OfficeCLI's dispatcher: a column break when a column is named.
        let column = props.iter().any(|(key, _)| {
            key.as_str().eq_ignore_ascii_case("col") || key.as_str().eq_ignore_ascii_case("column")
        });
        return Some((
            if column {
                Kind::ColBreak
            } else {
                Kind::RowBreak
            },
            None,
        ));
    }
    let canonical = match name.to_ascii_lowercase().as_str() {
        "img" => "picture",
        "textbox" => "shape",
        _ => crate::path::vocabulary::WORKBOOK.canonical(name)?,
    };
    let kind = <Kind as crate::path::navigate::ElementKind>::from_name(canonical)?;
    features::is_feature(kind).then_some((kind, None))
}

#[cfg(test)]
mod tests {
    use super::native;
    use crate::error::Code;
    use litchi_core::edit::{Refusal, RefusalKind};

    #[test]
    fn native_uses_typed_leaf_classification() {
        assert_eq!(
            native(litchi_xls::sheet_features::edit::Error::Refused(
                Refusal::new(RefusalKind::Protected, "set_note", "the workbook is signed",),
            ))
            .code(),
            Code::DocumentProtected
        );
        assert_eq!(
            native(litchi_xls::sheet_features::edit::Error::Refused(
                Refusal::new(
                    RefusalKind::InvalidValue,
                    "set_note",
                    "the target cell already has a note",
                ),
            ))
            .code(),
            Code::InvalidValue
        );
        assert_eq!(
            native(litchi_xls::sheet_features::edit::Error::PatchConflict).code(),
            Code::InvalidInput
        );
        assert_eq!(
            native(litchi_xls::sheet_features::edit::Error::Signed).code(),
            Code::DocumentProtected
        );
    }
}
