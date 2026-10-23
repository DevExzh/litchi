//! Add, set and remove of XLSB worksheet features (record 0322).
//!
//! Every verb admits its properties through the kind's registry table in
//! `features` (the XLSX vocabulary of record 0320), builds the leaf value,
//! and publishes one leaf commit of the family that owns the feature:
//!
//! | Kind | Leaf transaction |
//! |---|---|
//! | `comment` | `comments::edit` (notes and their VML note shapes) |
//! | `hyperlink` | `hyperlinks::edit` (`BrtHLink` and its relationship) |
//! | `picture`, `shape` | `drawing::edit` |
//! | `cf` | `conditional_formatting::edit` (with a `BrtDXF` for its formatting) |
//! | `dataValidation` | `data_validation::edit` |
//! | `autofilter` | `auto_filter::edit` |
//! | `rowbreak`, `colbreak` | `page_breaks::edit` |
//!
//! The leaf edits stage over the handle's workbook without copying it into
//! a package first. Property names and defaults follow OfficeCLI's Excel
//! handler; where this adapter is stricter (an unknown CF `type`, a list
//! literal with quotes) it refuses rather than guesses.

use std::sync::OnceLock;

use litchi_core::selector::{Path, Predicate, Step};
use litchi_core::{FileFormat, Position};
use litchi_xlsb as xlsb;
use xlsb::conditional_formatting::edit::{DifferentialStyle, Underline};
use xlsb::conditional_formatting::{Formatting, Rule, RuleType};
use xlsb::drawing::edit::{
    Align, COLUMN_EMU, Change, ElementKind, Fill, Font, Line, Marker, PictureSpec, Placement,
    ROW_EMU, ShapeSpec, VerticalAlign,
};

use super::features::{
    self, AutofilterKey, BreakKey, CfKey, CommentKey, Data, Feature, HyperlinkKey, PictureKey,
    ShapeKey, ValidationKey,
};
use super::node::{Item, Node};
use super::spec::Kind;
use super::write::{Commit, snapshot_limit};
use super::{Nav, Xlsb};
use crate::error::{Code, Error, Result};
use crate::new_node::NewNode;
use crate::placement::Placement as NodePlacement;
use crate::props::Props;
use crate::registry::{Operation, Table};
use crate::target::TargetArg;
use crate::transaction::SemanticLimits;
use crate::value::{Color, ColorContext, Length, Scalar};

fn error(code: Code, message: impl Into<String>) -> Error {
    Error::new(code, message.into()).with_format(FileFormat::Xlsb)
}

fn invalid(key: &'static str, message: impl Into<String>) -> Error {
    error(Code::InvalidValue, message).with_property(key)
}

/// A leaf refusal of authored input: an invalid value, not a corrupt file.
fn authored(error: xlsb::package::PackageError) -> Error {
    use xlsb::package::PackageError as Leaf;
    match error {
        Leaf::InvalidFormat(message)
        | Leaf::InvalidFormula(message)
        | Leaf::InvalidCellReference(message)
        | Leaf::Encoding(message) => {
            Error::new(Code::InvalidValue, message).with_format(FileFormat::Xlsb)
        },
        Leaf::Unrecognized { typ, val } => {
            Error::new(Code::InvalidValue, format!("{typ}: {val}")).with_format(FileFormat::Xlsb)
        },
        error => Error::from(error).with_format(FileFormat::Xlsb),
    }
}

fn native(error: xlsb::package::PackageError) -> Error {
    Error::from(error).with_format(FileFormat::Xlsb)
}

macro_rules! feature_edit_options {
    ($family:ident, $maximum:expr) => {
        xlsb::$family::edit::EditOptions::default().max_snapshot_bytes($maximum)
    };
}

// ---------------------------------------------------------------------------
// The transaction-history images of one feature commit

/// The canonical package images around one feature commit.
///
/// The images serialize on first use: a one-shot write never needs them,
/// and transaction history reads them once. A serialization failure yields
/// an image that equals no package image, so history reports a change
/// rather than hiding one.
#[derive(Debug, Clone)]
pub(crate) struct Images {
    before_file: Xlsb,
    after_file: Option<Xlsb>,
    before: OnceLock<Vec<u8>>,
    after: OnceLock<Vec<u8>>,
    /// The part payload bytes the leaf patch retains.
    retained: usize,
    /// The parts the leaf patch changed.
    parts: usize,
    /// The authored feature command count. Native part counts are zero for an
    /// exact no-op, but the unified leaf still records the admitted command.
    authored_operations: usize,
    /// The real native identity proof used when a signed private feature
    /// probe is discarded.  Unsigned feature commits intentionally keep this
    /// absent so publishing a normal no-op does not walk retained payloads.
    native: Option<NativeFeatureProof>,
}

impl Images {
    fn image(file: &Xlsb) -> Vec<u8> {
        let mut bytes = Vec::new();
        match file.write_to(&mut bytes) {
            Ok(()) => bytes,
            Err(_) => b"<unserializable feature image>".to_vec(),
        }
    }

    pub(super) fn before(&self) -> &[u8] {
        self.before.get_or_init(|| Self::image(&self.before_file))
    }

    pub(super) fn after(&self) -> &[u8] {
        match &self.after_file {
            Some(file) => self.after.get_or_init(|| Self::image(file)),
            None => self.before(),
        }
    }

    pub(super) const fn retained_bytes(&self) -> usize {
        self.retained
    }

    pub(super) fn is_noop(&self) -> bool {
        self.parts == 0 && self.native.as_ref().is_none_or(NativeFeatureProof::is_noop)
    }

    pub(super) const fn operation_count(&self) -> usize {
        self.authored_operations
    }
}

/// One of the native family patches behind the unified feature vocabulary.
/// The facade history still keeps package images for the heterogeneous leaf,
/// but a signed cancellation carries the actual native identity patch instead
/// of an inferred empty marker.
#[derive(Debug, Clone)]
enum NativeFeatureProof {
    Comments(xlsb::comments::edit::Patch),
    Hyperlinks(xlsb::hyperlinks::edit::Patch),
    Drawings(xlsb::drawing::edit::Patch),
    ConditionalFormatting(xlsb::conditional_formatting::edit::Patch),
    DataValidation(xlsb::data_validation::edit::Patch),
    AutoFilter(xlsb::auto_filter::edit::Patch),
    PageBreaks(xlsb::page_breaks::edit::Patch),
}

impl NativeFeatureProof {
    fn is_noop(&self) -> bool {
        match self {
            Self::Comments(patch) => patch.is_noop(),
            Self::Hyperlinks(patch) => patch.is_noop(),
            Self::Drawings(patch) => patch.is_noop(),
            Self::ConditionalFormatting(patch) => patch.is_noop(),
            Self::DataValidation(patch) => patch.is_noop(),
            Self::AutoFilter(patch) => patch.is_noop(),
            Self::PageBreaks(patch) => patch.is_noop(),
        }
    }
}

/// The common part of the seven native feature commits.  Keeping this
/// adapter typed means signed cancellation can rebuild the correct family
/// patch without serializing or fabricating a sentinel patch.
trait FeatureCommit: Sized {
    fn parts(&self) -> usize;
    fn retained_bytes(&self) -> usize;
    fn into_package(self) -> xlsb::Package;
    fn canonical_identity(
        base: &Xlsb,
        max_snapshot_bytes: usize,
    ) -> Result<(xlsb::Package, NativeFeatureProof, usize)>;
}

macro_rules! feature_commit {
    ($commit:path, $edit:ident, $edit_with_options:ident, $family:ident, $proof:ident) => {
        impl FeatureCommit for $commit {
            fn parts(&self) -> usize {
                self.diagnostics().parts()
            }

            fn retained_bytes(&self) -> usize {
                self.patch().retained_bytes()
            }

            fn into_package(self) -> xlsb::Package {
                <$commit>::into_package(self)
            }

            fn canonical_identity(
                base: &Xlsb,
                max_snapshot_bytes: usize,
            ) -> Result<(xlsb::Package, NativeFeatureProof, usize)> {
                let original = base.original_package_for_noop()?;
                let commit = original
                    .$edit_with_options(feature_edit_options!($family, max_snapshot_bytes))
                    .map_err(native)?
                    .commit()
                    .map_err(native)?;
                let (package, patch) = commit.into_parts();
                let retained = patch.retained_bytes();
                Ok((package, NativeFeatureProof::$proof(patch), retained))
            }
        }
    };
}

feature_commit!(
    xlsb::comments::edit::Commit,
    edit_comments,
    edit_comments_with_options,
    comments,
    Comments
);
feature_commit!(
    xlsb::hyperlinks::edit::Commit,
    edit_hyperlinks,
    edit_hyperlinks_with_options,
    hyperlinks,
    Hyperlinks
);
feature_commit!(
    xlsb::drawing::edit::Commit,
    edit_drawings,
    edit_drawings_with_options,
    drawing,
    Drawings
);
feature_commit!(
    xlsb::conditional_formatting::edit::Commit,
    edit_conditional_formats,
    edit_conditional_formats_with_options,
    conditional_formatting,
    ConditionalFormatting
);
feature_commit!(
    xlsb::data_validation::edit::Commit,
    edit_data_validations,
    edit_data_validations_with_options,
    data_validation,
    DataValidation
);
feature_commit!(
    xlsb::auto_filter::edit::Commit,
    edit_auto_filter,
    edit_auto_filter_with_options,
    auto_filter,
    AutoFilter
);
feature_commit!(
    xlsb::page_breaks::edit::Commit,
    edit_page_breaks,
    edit_page_breaks_with_options,
    page_breaks,
    PageBreaks
);

impl Xlsb {
    /// Publish one leaf feature commit: the new package (or this handle for
    /// an exact no-op) and the history images.
    fn publish<C: FeatureCommit>(
        &self,
        commit: C,
        authored_operations: usize,
        max_snapshot_bytes: usize,
    ) -> Result<Commit> {
        self.check_source_for_replay()?;
        let parts = commit.parts();
        let signed = self.signature() == super::Signature::Signed;
        if parts != 0 && signed {
            return Err(Error::signed(FileFormat::Xlsb));
        }
        let (package, native, retained) = if signed {
            let staged_retained = commit.retained_bytes();
            // The staged candidate is a private probe for a signed
            // cancellation.  Drop its package and native proof before
            // rebuilding the original-owner identity patch so the peak is
            // represented by the explicit max below, not by two live
            // endpoint graphs hidden behind that max.
            drop(commit);
            let (package, native, canonical_retained) =
                C::canonical_identity(self, max_snapshot_bytes)?;
            let retained = staged_retained.max(canonical_retained);
            if retained > max_snapshot_bytes {
                return Err(snapshot_limit(retained, max_snapshot_bytes));
            }
            (package, Some(native), retained)
        } else {
            // Native feature patches retain their changed part payloads even
            // for an unsigned publication.  Keep that cheap, already-cached
            // charge in the facade diagnostics so a snapshot limit cannot be
            // bypassed by the ordinary unsigned route.
            let retained = commit.retained_bytes();
            if retained > max_snapshot_bytes {
                return Err(snapshot_limit(retained, max_snapshot_bytes));
            }
            (commit.into_package(), None, retained)
        };
        let after_file = if parts == 0 {
            None
        } else {
            Some(self.adopt_package(package)?)
        };
        let after = after_file.clone().unwrap_or_else(|| self.clone());
        self.check_source_for_replay()?;
        Ok(Commit::from_feature(
            after,
            Images {
                before_file: self.clone(),
                after_file,
                before: OnceLock::new(),
                after: OnceLock::new(),
                retained,
                parts,
                authored_operations,
                native,
            },
        ))
    }
}

/// Split a leaf commit of any family into what [`Xlsb::publish`] needs.
macro_rules! publish {
    ($file:expr, $edit:expr, $limits:expr) => {{
        let commit = $edit.commit().map_err(native)?;
        $file.publish(commit, 1, $limits.max_snapshot_bytes())
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
    Length(Length),
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
                #[allow(clippy::cast_precision_loss, reason = "angles and sizes are small")]
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

/// A zero-based cell `(row, column)`.
type Cell = (u32, u32);

/// A zero-based inclusive rectangle `(first, last)`.
type Area = (Cell, Cell);

/// Parse an A1 cell (dollar signs ignored).
fn parse_cell(text: &str) -> Option<Cell> {
    let cell = litchi_core::selector::CellRef::parse(&text.trim().replace('$', ""))?;
    Some((cell.row(), cell.column()))
}

/// Parse an A1 cell or range (dollar signs ignored).
fn parse_area(text: &str) -> Option<Area> {
    let text = text.trim();
    match text.split_once(':') {
        Some((first, last)) => {
            let (first, last) = (parse_cell(first)?, parse_cell(last)?);
            (first.0 <= last.0 && first.1 <= last.1).then_some((first, last))
        },
        None => parse_cell(text).map(|cell| (cell, cell)),
    }
}

fn area_label(area: Area) -> String {
    crate::format::workbook::Label::range((area.0.1, area.0.0), (area.1.1, area.1.0))
        .as_str()
        .to_owned()
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
            reference: Some(features::a1(cell.at.0, cell.at.1)),
        }),
        Item::Range(range) => {
            let start = range.rect.start();
            let (rows, columns) = range.rect.end();
            Ok(Parent {
                sheet,
                reference: Some(area_label((
                    (start.row().get(), start.column().get()),
                    (rows.saturating_sub(1), columns.saturating_sub(1)),
                ))),
            })
        },
        _ => Err(unsupported_parent()),
    }
}

fn unsupported_parent() -> Error {
    error(
        Code::UnsupportedType,
        "worksheet features are added under a worksheet, a cell or a range",
    )
}

fn remove_at<T>(items: &mut Vec<T>, index: u32) -> Result<T> {
    let index = index as usize;
    if index >= items.len() {
        return Err(error(Code::PathNotFound, "the feature is missing"));
    }
    Ok(items.remove(index))
}

// ---------------------------------------------------------------------------
// Dispatch

impl Xlsb {
    /// The features of `kind` on the worksheet at `sheet`, read again.
    fn features_of(&self, sheet: usize, kind: Kind) -> Result<Vec<Feature>> {
        let handle = super::node::SheetHandle::at(self, sheet)?
            .ok_or_else(|| error(Code::PathNotFound, "the worksheet is missing"))?;
        self.finish(|| features::list_handle(self, &handle, kind))
    }

    /// Add a worksheet feature under `parent`; `implied` is the rule type a
    /// CF element word (`databar`, `cellis`) names.
    pub(super) fn add_feature(
        &self,
        parent_node: &Node<'_>,
        kind: Kind,
        implied: Option<&'static str>,
        node: &NewNode<'_>,
        max_snapshot_bytes: usize,
    ) -> Result<(Commit, Path<'static>)> {
        if !matches!(node.placement(), None | Some(NodePlacement::End)) {
            return Err(error(
                Code::UnsupportedCapability,
                "worksheet features are added at the end of their kind",
            )
            .with_suggestion("omit placement or use Placement::End"));
        }
        let parent = parent(parent_node)?;
        if !super::node::SheetHandle::at(self, parent.sheet)?.is_some_and(|handle| handle.is_grid())
        {
            return Err(error(
                Code::UnsupportedType,
                "worksheet features belong to worksheets",
            ));
        }
        match kind {
            Kind::Comment => self.add_comment(&parent, node.props(), max_snapshot_bytes),
            Kind::Hyperlink => self.add_hyperlink(&parent, node.props(), max_snapshot_bytes),
            Kind::Picture => self.add_picture(&parent, node.props(), max_snapshot_bytes),
            Kind::Shape => self.add_shape(&parent, node.props(), max_snapshot_bytes),
            Kind::Cf => self.add_cf(&parent, implied, node.props(), max_snapshot_bytes),
            Kind::DataValidation => self.add_validation(&parent, node.props(), max_snapshot_bytes),
            Kind::Autofilter => self.add_autofilter(&parent, node.props(), max_snapshot_bytes),
            Kind::RowBreak | Kind::ColBreak => {
                self.add_break(&parent, kind, node.props(), max_snapshot_bytes)
            },
            _ => Err(error(
                Code::UnsupportedElement,
                "this element kind cannot be added to an XLSB worksheet",
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
    ) -> Result<Option<Commit>> {
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
    ) -> Result<Option<Commit>> {
        let Some(node) = self.feature_target(target)? else {
            return Ok(None);
        };
        let Item::Feature(feature) = node.item() else {
            return Ok(None);
        };
        let sheet = node.sheet_position();
        let editor = self.editor_for_write()?;
        let commit = match &feature.data {
            Data::Comment(note) => {
                let mut edit = editor
                    .edit_comments_with_options(feature_edit_options!(
                        comments,
                        limits.max_snapshot_bytes()
                    ))
                    .map_err(native)?;
                edit.remove(sheet, note.row, note.col).map_err(native)?;
                publish!(self, edit, limits)
            },
            Data::Hyperlink(_) => {
                let mut edit = editor
                    .edit_hyperlinks_with_options(feature_edit_options!(
                        hyperlinks,
                        limits.max_snapshot_bytes()
                    ))
                    .map_err(native)?;
                edit.remove(sheet, feature.index as usize).map_err(native)?;
                publish!(self, edit, limits)
            },
            Data::Drawing(_) => {
                let kind = if feature.kind == Kind::Picture {
                    ElementKind::Picture
                } else {
                    ElementKind::Shape
                };
                let mut edit = editor
                    .edit_drawings_with_options(feature_edit_options!(
                        drawing,
                        limits.max_snapshot_bytes()
                    ))
                    .map_err(native)?;
                edit.remove(sheet, kind, feature.index as usize)
                    .map_err(native)?;
                publish!(self, edit, limits)
            },
            Data::Cf(_) => {
                let mut edit = editor
                    .edit_conditional_formats_with_options(feature_edit_options!(
                        conditional_formatting,
                        limits.max_snapshot_bytes()
                    ))
                    .map_err(native)?;
                edit.remove(sheet, feature.index as usize).map_err(native)?;
                publish!(self, edit, limits)
            },
            Data::Validation(_) => {
                let mut edit = editor
                    .edit_data_validations_with_options(feature_edit_options!(
                        data_validation,
                        limits.max_snapshot_bytes()
                    ))
                    .map_err(native)?;
                edit.remove(sheet, feature.index as usize).map_err(native)?;
                publish!(self, edit, limits)
            },
            Data::Autofilter(_) => {
                let mut edit = editor
                    .edit_auto_filter_with_options(feature_edit_options!(
                        auto_filter,
                        limits.max_snapshot_bytes()
                    ))
                    .map_err(native)?;
                edit.set_auto_filter(sheet, None).map_err(native)?;
                publish!(self, edit, limits)
            },
            Data::Break(_) => {
                let mut breaks = self.breaks(sheet, feature.kind)?;
                remove_at(&mut breaks, feature.index)?;
                return self
                    .publish_breaks(sheet, feature.kind, breaks, limits.max_snapshot_bytes())
                    .map(Some);
            },
        }?;
        Ok(Some(commit))
    }
}

// ---------------------------------------------------------------------------
// Notes

fn check_cell(key: &'static str, reference: &str) -> Result<Cell> {
    parse_cell(reference).ok_or_else(|| invalid(key, "the reference is not one A1 cell"))
}

impl Xlsb {
    fn add_comment(
        &self,
        parent: &Parent,
        props: &Props<'_>,
        max_snapshot_bytes: usize,
    ) -> Result<(Commit, Path<'static>)> {
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
            (None, None) => {
                return Err(error(Code::MissingProperty, "a note needs a ref").with_property("ref"));
            },
        };
        let editor = self.editor_for_write()?;
        let mut edit = editor
            .edit_comments_with_options(feature_edit_options!(comments, max_snapshot_bytes))
            .map_err(native)?;
        let index = edit
            .add(
                parent.sheet,
                row,
                col,
                author.as_deref().unwrap_or("Author"),
                text.as_deref().unwrap_or(""),
            )
            .map_err(authored)?;
        // Notes are numbered row-major, as `features::list` numbers them.
        let Some(index) = index else {
            return Err(error(Code::InvalidArgument, "the cell already has a note")
                .with_property("ref")
                .with_suggestion("set the existing comment instead"));
        };
        let commit = self.publish(edit.commit().map_err(native)?, 1, max_snapshot_bytes)?;
        Ok((commit, feature_path(parent.sheet, "comment", Some(index))))
    }

    fn set_comment(
        &self,
        sheet: usize,
        feature: &Feature,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        let Data::Comment(note) = &feature.data else {
            return Err(error(Code::InternalError, "a note node holds no note"));
        };
        let mut cell = (note.row, note.col);
        let mut author = note.author.clone();
        let mut text = note.text.clone();
        for (key, name, arg) in admit(&features::COMMENT_TABLE, Operation::Set, props)? {
            match key {
                CommentKey::Ref => cell = check_cell(name, &arg.text().unwrap_or_default())?,
                CommentKey::Author => author = arg.text().unwrap_or_default(),
                CommentKey::Text => text = arg.text().unwrap_or_default(),
                CommentKey::AnchoredTo => {},
            }
        }
        let editor = self.editor_for_write()?;
        let mut edit = editor
            .edit_comments_with_options(feature_edit_options!(
                comments,
                limits.max_snapshot_bytes()
            ))
            .map_err(native)?;
        if cell != (note.row, note.col) {
            if edit
                .comments(sheet)
                .map_err(native)?
                .iter()
                .any(|other| (other.row, other.col) == cell)
            {
                return Err(
                    error(Code::InvalidArgument, "the target cell already has a note")
                        .with_property("ref"),
                );
            }
            edit.remove(sheet, note.row, note.col).map_err(native)?;
        }
        edit.set(sheet, cell.0, cell.1, &author, &text)
            .map_err(authored)?;
        publish!(self, edit, limits)
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
) -> Result<xlsb::hyperlinks::Hyperlink> {
    let ((row_first, col_first), (row_last, col_last)) =
        parse_area(reference).ok_or_else(|| invalid("ref", "the reference is not an A1 range"))?;
    if location.is_none() && url.is_none() {
        return Err(error(
            Code::MissingProperty,
            "a hyperlink needs a url or a location",
        )
        .with_property("url"));
    }
    let mut link =
        xlsb::hyperlinks::Hyperlink::new(row_first, row_last, col_first, col_last, String::new());
    link.location = location;
    link.target = url;
    link.display = display;
    link.tooltip = tooltip;
    Ok(link)
}

impl Xlsb {
    fn add_hyperlink(
        &self,
        parent: &Parent,
        props: &Props<'_>,
        max_snapshot_bytes: usize,
    ) -> Result<(Commit, Path<'static>)> {
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
        let reference = reference.ok_or_else(|| {
            error(Code::MissingProperty, "a hyperlink needs a ref").with_property("ref")
        })?;
        let link = hyperlink(&reference, location, url, display, tooltip)?;
        let editor = self.editor_for_write()?;
        let mut edit = editor
            .edit_hyperlinks_with_options(feature_edit_options!(hyperlinks, max_snapshot_bytes))
            .map_err(native)?;
        // As XLSX's `put_hyperlink`: a link on the same cells is replaced.
        let existing = edit
            .hyperlinks(parent.sheet)
            .map_err(native)?
            .iter()
            .position(|held| features::link_reference(held) == features::link_reference(&link));
        let index = match existing {
            Some(index) => {
                edit.set(parent.sheet, index, link).map_err(authored)?;
                index
            },
            None => edit.add(parent.sheet, link).map_err(authored)?,
        };
        let commit = self.publish(edit.commit().map_err(native)?, 1, max_snapshot_bytes)?;
        Ok((commit, feature_path(parent.sheet, "hyperlink", Some(index))))
    }

    fn set_hyperlink(
        &self,
        sheet: usize,
        feature: &Feature,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        let Data::Hyperlink(link) = &feature.data else {
            return Err(error(
                Code::InternalError,
                "a hyperlink node holds no hyperlink",
            ));
        };
        let mut reference = features::link_reference(link);
        let mut location = link.location.clone();
        let mut url = link.target.clone();
        let mut display = link.display.clone();
        let mut tooltip = link.tooltip.clone();
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
        let editor = self.editor_for_write()?;
        let mut edit = editor
            .edit_hyperlinks_with_options(feature_edit_options!(
                hyperlinks,
                limits.max_snapshot_bytes()
            ))
            .map_err(native)?;
        edit.set(sheet, feature.index as usize, replacement)
            .map_err(authored)?;
        publish!(self, edit, limits)
    }
}

// ---------------------------------------------------------------------------
// Pictures and shapes

/// A picture source: a base64 `data:` URI or a local file path.
fn read_source(src: &str) -> Result<Vec<u8>> {
    const MAXIMUM: usize = xlsb::drawing::edit::MAX_IMAGE_BYTES;
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
        if payload.len() / 4 * 3 > MAXIMUM {
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
    if length > MAXIMUM as u64 {
        return Err(invalid("src", "the image is too large"));
    }
    std::fs::read(src).map_err(|source| Error::from(source).with_property("src"))
}

/// The image content type and file extension of `data`, by signature.
fn classify(data: &[u8]) -> Option<(&'static str, &'static str)> {
    if data.starts_with(b"\x89PNG\r\n\x1a\n") {
        Some(("image/png", "png"))
    } else if data.starts_with(&[0xFF, 0xD8, 0xFF]) {
        Some(("image/jpeg", "jpeg"))
    } else if data.starts_with(b"GIF87a") || data.starts_with(b"GIF89a") {
        Some(("image/gif", "gif"))
    } else if data.starts_with(b"BM") {
        Some(("image/bmp", "bmp"))
    } else if data.starts_with(b"II*\0") || data.starts_with(b"MM\0*") {
        Some(("image/tiff", "tiff"))
    } else {
        None
    }
}

/// A geometry value: a count of cells, or a length.
#[derive(Clone, Copy)]
enum Measure {
    Cells(u32),
    Emu(i64),
}

fn measure(key: &'static str, arg: &Arg) -> Result<Measure> {
    let cells = |value: i64| {
        u32::try_from(value)
            .ok()
            .filter(|value| *value <= 1_048_575)
            .map(Measure::Cells)
            .ok_or_else(|| invalid(key, "the cell count is out of range"))
    };
    match arg {
        Arg::Int(value) => cells(*value),
        Arg::Text(value) => {
            let value = value.trim();
            if !value.is_empty() && value.bytes().all(|byte| byte.is_ascii_digit()) {
                return cells(
                    value
                        .parse()
                        .map_err(|_| invalid(key, "the value is out of range"))?,
                );
            }
            Length::parse(value)
                .map(|length| Measure::Emu(length.emu()))
                .map_err(|_| invalid(key, "the value is a cell count or a length with a unit"))
        },
        Arg::Length(length) => Ok(Measure::Emu(length.emu())),
        _ => Err(invalid(
            key,
            "the value is a cell count or a length with a unit",
        )),
    }
}

fn emu(measure: Measure, unit: i64) -> i64 {
    match measure {
        Measure::Cells(cells) => i64::from(cells) * unit,
        Measure::Emu(emu) => emu,
    }
}

/// The geometry properties of a picture or shape request.
#[derive(Default)]
struct Geometry {
    x: Option<Measure>,
    y: Option<Measure>,
    width: Option<Measure>,
    height: Option<Measure>,
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

    /// The placement, from `current` (a set) or defaults (an add).
    fn placement(&self, current: Option<Placement>, defaults: [u32; 4]) -> Result<Placement> {
        if let Some(anchor) = &self.anchor {
            let anchor = anchor.trim().replace('$', "");
            if !anchor.contains(':') {
                // A single cell spans four columns and two rows.
                let (row, col) = parse_cell(&anchor)
                    .ok_or_else(|| invalid("anchor", "the anchor is not an A1 range"))?;
                return Ok(Placement::TwoCell {
                    from: Marker::cell(col, row),
                    to: Marker::cell(col + 4, row + 2),
                });
            }
            let (first, last) = parse_area(&anchor)
                .ok_or_else(|| invalid("anchor", "the anchor is not an A1 range"))?;
            return Ok(Placement::TwoCell {
                from: Marker::cell(first.1, first.0),
                to: Marker::cell(last.1 + 1, last.0 + 1),
            });
        }
        let (mut x, mut y, mut width, mut height) = match current {
            Some(Placement::TwoCell { from, to }) => (
                Measure::Cells(from.col),
                Measure::Cells(from.row),
                Measure::Cells(to.col.saturating_sub(from.col)),
                Measure::Cells(to.row.saturating_sub(from.row)),
            ),
            Some(Placement::OneCell { from, cx, cy }) => (
                Measure::Cells(from.col),
                Measure::Cells(from.row),
                Measure::Emu(cx),
                Measure::Emu(cy),
            ),
            Some(Placement::Absolute { x, y, cx, cy }) => (
                Measure::Emu(x),
                Measure::Emu(y),
                Measure::Emu(cx),
                Measure::Emu(cy),
            ),
            _ => (
                Measure::Cells(defaults[0]),
                Measure::Cells(defaults[1]),
                Measure::Cells(defaults[2]),
                Measure::Cells(defaults[3]),
            ),
        };
        x = self.x.unwrap_or(x);
        y = self.y.unwrap_or(y);
        width = self.width.unwrap_or(width);
        height = self.height.unwrap_or(height);
        let mode = self
            .mode
            .as_deref()
            .map(str::to_ascii_lowercase)
            .or_else(|| match current {
                Some(Placement::OneCell { .. }) => Some("onecell".to_owned()),
                Some(Placement::Absolute { .. }) => Some("absolute".to_owned()),
                _ => None,
            });
        let all_cells = |measures: [Measure; 4]| {
            measures
                .iter()
                .all(|measure| matches!(measure, Measure::Cells(_)))
        };
        match mode.as_deref() {
            Some("absolute") => Ok(Placement::Absolute {
                x: emu(x, COLUMN_EMU),
                y: emu(y, ROW_EMU),
                cx: emu(width, COLUMN_EMU),
                cy: emu(height, ROW_EMU),
            }),
            Some("onecell") => {
                let (Measure::Cells(col), Measure::Cells(row)) = (x, y) else {
                    return Err(invalid("x", "a one-cell anchor's origin is a cell"));
                };
                Ok(Placement::OneCell {
                    from: Marker::cell(col, row),
                    cx: emu(width, COLUMN_EMU),
                    cy: emu(height, ROW_EMU),
                })
            },
            Some("twocell") | None if all_cells([x, y, width, height]) => {
                let cells = |measure: Measure| match measure {
                    Measure::Cells(cells) => cells,
                    Measure::Emu(_) => 0,
                };
                let (col, row) = (cells(x), cells(y));
                if cells(width) == 0 || cells(height) == 0 {
                    return Err(invalid("width", "a size must be positive"));
                }
                Ok(Placement::TwoCell {
                    from: Marker::cell(col, row),
                    to: Marker::cell(
                        col.saturating_add(cells(width)),
                        row.saturating_add(cells(height)),
                    ),
                })
            },
            None => {
                let (Measure::Cells(col), Measure::Cells(row)) = (x, y) else {
                    return Err(invalid("x", "a picture's origin is a cell"));
                };
                Ok(Placement::OneCell {
                    from: Marker::cell(col, row),
                    cx: emu(width, COLUMN_EMU),
                    cy: emu(height, ROW_EMU),
                })
            },
            Some(_) => Err(invalid(
                "anchorMode",
                "anchorMode is oneCell, twoCell or absolute",
            )),
        }
    }
}

/// A text margin: one length (or points) for every side, or four
/// comma-separated values (left, top, right, bottom).
fn margin(key: &'static str, value: &str) -> Result<[i64; 4]> {
    let inset = |text: &str| -> Result<i64> {
        let text = text.trim();
        match text.parse::<f64>() {
            Ok(points) if points.is_finite() && points >= 0.0 => {
                #[allow(clippy::cast_possible_truncation, reason = "bounded by the leaf")]
                let emu = (points * 12_700.0).round() as i64;
                Ok(emu)
            },
            _ => Length::parse(text)
                .map(Length::emu)
                .map_err(|_| invalid(key, "a margin is a number of points or a length")),
        }
    };
    let parts: Vec<&str> = value.split(',').collect();
    Ok(match parts.as_slice() {
        [all] => [inset(all)?; 4],
        [left, top, right, bottom] => [inset(left)?, inset(top)?, inset(right)?, inset(bottom)?],
        _ => return Err(invalid(key, "a margin is one value or four")),
    })
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

impl Xlsb {
    fn drawing_count(&self, sheet: usize, kind: Kind) -> Result<usize> {
        Ok(self.features_of(sheet, kind)?.len())
    }

    fn add_picture(
        &self,
        parent: &Parent,
        props: &Props<'_>,
        max_snapshot_bytes: usize,
    ) -> Result<(Commit, Path<'static>)> {
        let mut src = None;
        let mut name = None;
        let mut description = None;
        let mut title = None;
        let mut rotation = None;
        let (mut flip_h, mut flip_v) = (false, false);
        let mut geometry = Geometry::default();
        for (key, name_key, arg) in admit(&features::PICTURE_TABLE, Operation::Add, props)? {
            match key {
                PictureKey::Src => src = arg.text(),
                PictureKey::Name => name = arg.text(),
                PictureKey::Alt => description = arg.text(),
                PictureKey::Title => title = arg.text(),
                PictureKey::Rotation => rotation = Some(arg.number(name_key)?),
                PictureKey::Flip => {
                    (flip_h, flip_v) = flip_value(name_key, &arg.text().unwrap_or_default())?
                },
                PictureKey::FlipH => flip_h = arg.flag(),
                PictureKey::FlipV => flip_v = arg.flag(),
                PictureKey::FlipBoth => (flip_h, flip_v) = (arg.flag(), arg.flag()),
                PictureKey::X => geometry.x = Some(measure(name_key, &arg)?),
                PictureKey::Y => geometry.y = Some(measure(name_key, &arg)?),
                PictureKey::Width => geometry.width = Some(measure(name_key, &arg)?),
                PictureKey::Height => geometry.height = Some(measure(name_key, &arg)?),
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
                PictureKey::Decorative
                | PictureKey::Hyperlink
                | PictureKey::ContentType
                | PictureKey::FileSize
                | PictureKey::Pending => {},
            }
        }
        let src = src.ok_or_else(|| {
            error(Code::MissingProperty, "a picture needs a src").with_property("src")
        })?;
        let image = read_source(&src)?;
        let (content_type, extension) = classify(&image).ok_or_else(|| {
            error(
                Code::UnsupportedCapability,
                "the picture is not a PNG, JPEG, GIF, BMP or TIFF image",
            )
            .with_property("src")
        })?;
        let placement = if geometry.mode.is_none() && geometry.anchor.is_none() {
            geometry.mode = Some("twocell".to_owned());
            geometry.placement(None, [0, 0, 5, 5])?
        } else {
            geometry.placement(None, [0, 0, 5, 5])?
        };
        let name = match name {
            Some(name) => name,
            None => format!(
                "Picture {}",
                self.drawing_count(parent.sheet, Kind::Picture)? + 1
            ),
        };
        let spec = PictureSpec {
            placement,
            name,
            description,
            title,
            image,
            content_type: content_type.to_owned(),
            extension: extension.to_owned(),
            rotation,
            flip_h,
            flip_v,
        };
        let editor = self.editor_for_write()?;
        let mut edit = editor
            .edit_drawings_with_options(feature_edit_options!(drawing, max_snapshot_bytes))
            .map_err(native)?;
        let index = edit.add_picture(parent.sheet, &spec).map_err(authored)?;
        let commit = self.publish(edit.commit().map_err(native)?, 1, max_snapshot_bytes)?;
        Ok((commit, feature_path(parent.sheet, "picture", Some(index))))
    }

    fn add_shape(
        &self,
        parent: &Parent,
        props: &Props<'_>,
        max_snapshot_bytes: usize,
    ) -> Result<(Commit, Path<'static>)> {
        let mut spec = ShapeSpec::new(
            Placement::TwoCell {
                from: Marker::cell(1, 1),
                to: Marker::cell(6, 4),
            },
            String::new(),
        );
        let mut name = None;
        let mut geometry = Geometry::default();
        for (key, name_key, arg) in admit(&features::SHAPE_TABLE, Operation::Add, props)? {
            let text = || arg.text().unwrap_or_default();
            match key {
                ShapeKey::Name => name = arg.text(),
                ShapeKey::Text => spec.text = Some(text().replace("\\n", "\n")),
                ShapeKey::Geometry => spec.geometry = text(),
                ShapeKey::X => geometry.x = Some(measure(name_key, &arg)?),
                ShapeKey::Y => geometry.y = Some(measure(name_key, &arg)?),
                ShapeKey::Width => geometry.width = Some(measure(name_key, &arg)?),
                ShapeKey::Height => geometry.height = Some(measure(name_key, &arg)?),
                ShapeKey::Anchor => geometry.anchor = arg.text(),
                ShapeKey::Fill => {
                    let value = text();
                    spec.fill = Some(if value.eq_ignore_ascii_case("none") {
                        Fill::None
                    } else if let Some(name) = xlsb::drawing::edit::canonical(
                        xlsb::drawing::edit::SCHEME_COLORS,
                        value.trim(),
                    ) {
                        Fill::Scheme(name)
                    } else {
                        Fill::Solid(parse_color(name_key, &value, ColorContext::Drawing)?)
                    });
                },
                ShapeKey::Line => {
                    let value = text();
                    spec.line = Some(if value.eq_ignore_ascii_case("none") {
                        Line::None
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
                                    .filter(|points| points.is_finite() && *points >= 0.0)
                                    .map(|points| {
                                        #[allow(
                                            clippy::cast_possible_truncation,
                                            reason = "bounded line width"
                                        )]
                                        let emu = (points * 12_700.0).round() as i64;
                                        emu
                                    })
                                    .ok_or_else(|| {
                                        invalid(name_key, "the line width is a number of points")
                                    })
                            })
                            .transpose()?;
                        let dash = parts
                            .next()
                            .map(|dash| {
                                xlsb::drawing::edit::canonical(
                                    xlsb::drawing::edit::DASHES,
                                    dash.trim(),
                                )
                                .ok_or_else(|| invalid(name_key, "the line style is a preset dash"))
                            })
                            .transpose()?;
                        Line::Solid { color, width, dash }
                    });
                },
                ShapeKey::Font => spec.font.name = arg.text(),
                ShapeKey::Size => {
                    if let Arg::Length(length) = arg {
                        spec.font.size = Some(length.points());
                    }
                },
                ShapeKey::Bold => spec.font.bold = Some(arg.flag()),
                ShapeKey::Italic => spec.font.italic = Some(arg.flag()),
                ShapeKey::Underline => {
                    spec.font.underline = match text().to_ascii_lowercase().as_str() {
                        "single" | "true" | "sng" => Some("sng".to_owned()),
                        "double" | "dbl" => Some("dbl".to_owned()),
                        "none" | "false" | "" => None,
                        _ => return Err(invalid(name_key, "underline is single, double or none")),
                    };
                },
                ShapeKey::Color => {
                    spec.font.color = Some(arg_color(name_key, &arg, ColorContext::Drawing)?)
                },
                ShapeKey::Align => {
                    spec.align = Some(match text().to_ascii_lowercase().as_str() {
                        "left" | "l" => Align::Left,
                        "center" | "centre" | "ctr" => Align::Center,
                        "right" | "r" => Align::Right,
                        "justify" | "just" => Align::Justify,
                        _ => {
                            return Err(invalid(
                                name_key,
                                "align is left, center, right or justify",
                            ));
                        },
                    });
                },
                ShapeKey::Valign => {
                    spec.valign = Some(match text().to_ascii_lowercase().as_str() {
                        "top" | "t" => VerticalAlign::Top,
                        "middle" | "center" | "ctr" => VerticalAlign::Middle,
                        "bottom" | "b" => VerticalAlign::Bottom,
                        _ => return Err(invalid(name_key, "valign is top, middle or bottom")),
                    });
                },
                ShapeKey::Rotation => spec.rotation = Some(arg.number(name_key)?),
                ShapeKey::Flip => (spec.flip_h, spec.flip_v) = flip_value(name_key, &text())?,
                ShapeKey::FlipH => spec.flip_h = arg.flag(),
                ShapeKey::FlipV => spec.flip_v = arg.flag(),
                ShapeKey::FlipBoth => (spec.flip_h, spec.flip_v) = (arg.flag(), arg.flag()),
                ShapeKey::Margin => spec.margin = Some(margin(name_key, &text())?),
                ShapeKey::Alt => spec.description = arg.text(),
                ShapeKey::Hyperlink | ShapeKey::Pending => {},
            }
        }
        if !geometry.is_empty() {
            if geometry.mode.is_none() {
                geometry.mode = Some("twocell".to_owned());
            }
            spec.placement = geometry.placement(None, [1, 1, 5, 3])?;
        }
        spec.name = match name {
            Some(name) => name,
            None => format!(
                "Shape {}",
                self.drawing_count(parent.sheet, Kind::Shape)? + 1
            ),
        };
        let _: &Font = &spec.font;
        let editor = self.editor_for_write()?;
        let mut edit = editor
            .edit_drawings_with_options(feature_edit_options!(drawing, max_snapshot_bytes))
            .map_err(native)?;
        let index = edit.add_shape(parent.sheet, &spec).map_err(authored)?;
        let commit = self.publish(edit.commit().map_err(native)?, 1, max_snapshot_bytes)?;
        Ok((commit, feature_path(parent.sheet, "shape", Some(index))))
    }

    fn set_drawing(
        &self,
        sheet: usize,
        feature: &Feature,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        let Data::Drawing(element) = &feature.data else {
            return Err(error(
                Code::InternalError,
                "a drawing node holds no element",
            ));
        };
        let mut change = Change::default();
        let mut geometry = Geometry::default();
        let (mut flip_h, mut flip_v) = (None, None);
        let picture = feature.kind == Kind::Picture;
        let admitted: Vec<(Option<PictureKey>, Option<ShapeKey>, &'static str, Arg)> = if picture {
            admit(&features::PICTURE_TABLE, Operation::Set, props)?
                .into_iter()
                .map(|(key, name, arg)| (Some(key), None, name, arg))
                .collect()
        } else {
            admit(&features::SHAPE_TABLE, Operation::Set, props)?
                .into_iter()
                .map(|(key, name, arg)| (None, Some(key), name, arg))
                .collect()
        };
        for (picture_key, shape_key, name, arg) in admitted {
            let field = match (picture_key, shape_key) {
                (Some(PictureKey::Name), _) | (_, Some(ShapeKey::Name)) => "name",
                (Some(PictureKey::Alt), _) | (_, Some(ShapeKey::Alt)) => "alt",
                (Some(PictureKey::X), _) | (_, Some(ShapeKey::X)) => "x",
                (Some(PictureKey::Y), _) | (_, Some(ShapeKey::Y)) => "y",
                (Some(PictureKey::Width), _) | (_, Some(ShapeKey::Width)) => "width",
                (Some(PictureKey::Height), _) | (_, Some(ShapeKey::Height)) => "height",
                (Some(PictureKey::Rotation), _) | (_, Some(ShapeKey::Rotation)) => "rotation",
                (Some(PictureKey::Flip), _) | (_, Some(ShapeKey::Flip)) => "flip",
                (Some(PictureKey::FlipH), _) | (_, Some(ShapeKey::FlipH)) => "flipH",
                (Some(PictureKey::FlipV), _) | (_, Some(ShapeKey::FlipV)) => "flipV",
                (Some(PictureKey::FlipBoth), _) | (_, Some(ShapeKey::FlipBoth)) => "flipBoth",
                _ => {
                    return Err(
                        error(Code::UnsupportedProperty, "the property is not settable")
                            .with_property(name),
                    );
                },
            };
            match field {
                "name" => change.name = arg.text(),
                "alt" => change.description = Some(arg.text()),
                "x" => geometry.x = Some(measure(name, &arg)?),
                "y" => geometry.y = Some(measure(name, &arg)?),
                "width" => geometry.width = Some(measure(name, &arg)?),
                "height" => geometry.height = Some(measure(name, &arg)?),
                "rotation" => change.rotation = Some(arg.number(name)?),
                "flip" => {
                    let (h, v) = flip_value(name, &arg.text().unwrap_or_default())?;
                    flip_h = Some(h);
                    flip_v = Some(v);
                },
                "flipH" => flip_h = Some(arg.flag()),
                "flipBoth" => {
                    flip_h = Some(arg.flag());
                    flip_v = Some(arg.flag());
                },
                _ => flip_v = Some(arg.flag()),
            }
        }
        change.flip_h = flip_h;
        change.flip_v = flip_v;
        if !geometry.is_empty() {
            change.placement = Some(geometry.placement(element.placement, [0, 0, 5, 5])?);
        }
        let kind = if picture {
            ElementKind::Picture
        } else {
            ElementKind::Shape
        };
        let editor = self.editor_for_write()?;
        let mut edit = editor
            .edit_drawings_with_options(feature_edit_options!(drawing, limits.max_snapshot_bytes()))
            .map_err(native)?;
        edit.set(sheet, kind, feature.index as usize, &change)
            .map_err(authored)?;
        publish!(self, edit, limits)
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
    /// A text rule with its `CFTextOper` (0 contains, 1 not contains, 2
    /// begins with, 3 ends with).
    Text(u32),
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
fn rule_kind(rule: &Rule) -> Option<CfKind> {
    Some(match rule.rule_type {
        RuleType::CellIs => CfKind::CellIs,
        RuleType::ColorScale => CfKind::ColorScale,
        RuleType::DataBar => CfKind::DataBar,
        RuleType::IconSet => CfKind::IconSet,
        RuleType::TopN => CfKind::TopN,
        RuleType::Expression => match rule.template {
            1 => CfKind::Formula,
            7 => CfKind::UniqueValues,
            8 => CfKind::Text(rule.parameter),
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

/// A cellIs comparison as its `CFOper` wire value.
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
fn cf_period(key: &'static str, value: &str) -> Result<u32> {
    let value = value.trim();
    features::PERIODS
        .iter()
        .position(|(name, _)| name.eq_ignore_ascii_case(value))
        .and_then(|index| u32::try_from(index).ok())
        .map(|index| index + 15)
        .ok_or_else(|| invalid(key, "the period is not a time period"))
}

fn cf_ranges(key: &'static str, value: &str) -> Result<Vec<String>> {
    let ranges: Vec<String> = value
        .split_whitespace()
        .map(|range| {
            parse_area(range)
                .map(area_label)
                .ok_or_else(|| invalid(key, "the range is not an A1 range"))
        })
        .collect::<Result<_>>()?;
    if ranges.is_empty() {
        return Err(invalid(key, "the range is empty"));
    }
    Ok(ranges)
}

/// The top-left cell of the first range, the anchor of generated formulas.
fn first_cell(ranges: &[String]) -> String {
    let range = ranges.first().map_or("A1", String::as_str);
    range.split(':').next().unwrap_or("A1").to_owned()
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
fn period_formula(template: u32, cell: &str) -> String {
    match template {
        15 => format!("FLOOR({cell},1)=TODAY()"),
        16 => format!("FLOOR({cell},1)=TODAY()+1"),
        17 => format!("FLOOR({cell},1)=TODAY()-1"),
        18 => format!("AND(TODAY()-FLOOR({cell},1)<=6,FLOOR({cell},1)<=TODAY())"),
        19 => format!(
            "AND(MONTH({cell})=MONTH(EDATE(TODAY(),0-1)),YEAR({cell})=YEAR(EDATE(TODAY(),0-1)))"
        ),
        20 => format!(
            "AND(MONTH({cell})=MONTH(EDATE(TODAY(),0+1)),YEAR({cell})=YEAR(EDATE(TODAY(),0+1)))"
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

/// A threshold value object; `fallback` is the `CFVOtype` an absent or
/// `auto` value takes.
fn cf_threshold(
    key: &'static str,
    value: Option<&str>,
    fallback: u8,
) -> Result<xlsb::conditional_formatting::Value> {
    use xlsb::conditional_formatting::Value as Threshold;
    let value = value
        .map(str::trim)
        .filter(|value| !value.is_empty() && !value.eq_ignore_ascii_case("auto"));
    let Some(value) = value else {
        return Ok(Threshold::new(fallback, None));
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
    Ok(Threshold::new(kind, Some(text.trim().to_owned())))
}

const fn argb(rgb: u32) -> u32 {
    0xFF00_0000 | (rgb & 0x00FF_FFFF)
}

/// The parts of a CF request, applied to a fresh or an existing rule.
#[derive(Default)]
struct CfRequest {
    kind: Option<CfType>,
    ranges: Option<Vec<String>>,
    stop_if_true: Option<bool>,
    operator: Option<u8>,
    value: Option<String>,
    value2: Option<String>,
    formula: Option<String>,
    text: Option<String>,
    rank: Option<u32>,
    percent: Option<bool>,
    bottom: Option<bool>,
    above_average: Option<bool>,
    equal_average: Option<bool>,
    std_dev: Option<u32>,
    period: Option<u32>,
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
    style: DifferentialStyle,
}

fn style_is_empty(style: &DifferentialStyle) -> bool {
    *style == DifferentialStyle::default()
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
            CfKey::Ref => request.ranges = Some(cf_ranges(name, &text())?),
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
                    u32::try_from(arg.int(name)?)
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
                    u32::try_from(arg.int(name)?)
                        .ok()
                        .filter(|value| (1..=3).contains(value))
                        .ok_or_else(|| invalid(name, "stdDev is 1, 2 or 3"))?,
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
                        .ok_or_else(|| invalid(name, "the icon set is not a core icon set"))?,
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
                    "single" | "true" => Some(Underline::Single),
                    "double" => Some(Underline::Double),
                    "none" | "false" => None,
                    _ => return Err(invalid(name, "underline is single, double or none")),
                };
            },
            CfKey::FontSize => {
                if let Arg::Length(length) = arg {
                    request.style.size = Some(length.points());
                }
            },
            CfKey::FontName => request.style.name = arg.text(),
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

fn missing(key: &'static str, message: &'static str) -> Error {
    error(Code::MissingProperty, message).with_property(key)
}

/// A fresh rule of `kind` with its kind's template.
fn fresh_rule(kind: CfKind, priority: u32) -> Rule {
    let rule_type = match kind {
        CfKind::CellIs => RuleType::CellIs,
        CfKind::ColorScale => RuleType::ColorScale,
        CfKind::DataBar => RuleType::DataBar,
        CfKind::IconSet => RuleType::IconSet,
        CfKind::TopN => RuleType::TopN,
        _ => RuleType::Expression,
    };
    let mut rule = Rule::new(rule_type, priority);
    rule.template = match kind {
        CfKind::Formula => 1,
        CfKind::UniqueValues => 7,
        CfKind::Text(_) => 8,
        CfKind::ContainsBlanks => 9,
        CfKind::NotContainsBlanks => 10,
        CfKind::ContainsErrors => 11,
        CfKind::NotContainsErrors => 12,
        CfKind::TimePeriod => 15,
        CfKind::AboveAverage => 25,
        CfKind::DuplicateValues => 27,
        _ => rule.template,
    };
    if let CfKind::Text(operator) = kind {
        rule.parameter = operator;
    }
    if kind == CfKind::AboveAverage {
        rule.above_average = true;
    }
    rule
}

/// Replace a rule's formulas with text the leaf compiles.
fn set_formulas(rule: &mut Rule, formulas: Vec<String>) {
    rule.formulas.clear();
    rule.formula_extras.clear();
    rule.formula_texts = formulas;
}

/// Apply a request to `rule`: a fresh rule of the request's kind (an add)
/// or the existing first rule of a format (a set).
fn apply_cf(rule: &mut Rule, request: &CfRequest, ranges: &[String], fresh: bool) -> Result<()> {
    let kind = rule_kind(rule).ok_or_else(|| {
        error(
            Code::UnsupportedCapability,
            "the conditional-format rule template is not one the XLSB editor authors",
        )
    })?;
    if let Some(value) = request.stop_if_true {
        rule.stop_if_true = value;
    }
    let moved = request.ranges.is_some();
    match kind {
        CfKind::CellIs => {
            if let Some(operator) = request.operator {
                rule.operator = Some(operator);
                rule.parameter = u32::from(operator);
            } else if fresh {
                rule.operator = Some(5);
                rule.parameter = 5;
            }
            let mut formulas = rule.formula_texts.clone();
            let changed = request.value.is_some() || request.value2.is_some() || moved;
            if let Some(value) = &request.value {
                if formulas.is_empty() {
                    formulas.push(value.clone());
                } else {
                    formulas[0].clone_from(value);
                }
            }
            if let Some(value) = &request.value2 {
                formulas.truncate(1);
                formulas.push(value.clone());
            }
            let between = matches!(rule.operator, Some(1 | 2));
            if !between {
                formulas.truncate(1);
            }
            if formulas.is_empty() {
                return Err(missing("value", "a cellIs rule needs a value"));
            }
            if between && formulas.len() < 2 {
                return Err(missing("value2", "a between rule needs value2"));
            }
            if changed || request.operator.is_some() || fresh {
                set_formulas(rule, formulas);
            }
        },
        CfKind::Formula => {
            if let Some(formula) = request.formula.clone().or_else(|| request.value.clone()) {
                set_formulas(rule, vec![formula]);
            } else if moved {
                let texts = rule.formula_texts.clone();
                set_formulas(rule, texts);
            }
            if rule.formula_texts.is_empty() && rule.formulas.is_empty() {
                return Err(missing("formula", "a formula rule needs a formula"));
            }
        },
        CfKind::TopN => {
            if let Some(rank) = request
                .rank
                .or_else(|| request.value.as_deref().and_then(|v| v.parse().ok()))
            {
                rule.parameter = rank;
            } else if fresh {
                rule.parameter = 10;
            }
            if let Some(value) = request.percent {
                rule.percent = value;
            }
            if let Some(value) = request.bottom {
                rule.bottom = value;
            }
        },
        CfKind::AboveAverage => {
            let below = !rule.above_average;
            let equal = matches!(rule.template, 29 | 30);
            let below = request.above_average.map_or(below, |above| !above);
            if let Some(value) = request.std_dev {
                rule.parameter = value;
            }
            // `CFTemp` has no equal-average template with standard
            // deviations: a deviation rule is strictly above or below.
            let equal = request.equal_average.unwrap_or(equal) && rule.parameter == 0;
            rule.template = match (below, equal) {
                (false, false) => 25,
                (true, false) => 26,
                (false, true) => 29,
                (true, true) => 30,
            };
            rule.above_average = !below;
        },
        CfKind::Text(_) => {
            if let Some(text) = &request.text {
                rule.text = Some(text.clone());
            }
            let text = rule
                .text
                .clone()
                .ok_or_else(|| missing("text", "a text rule needs text"))?;
            if fresh || request.text.is_some() || moved {
                let formula = generated_formula(kind, &text, &first_cell(ranges))
                    .ok_or_else(|| error(Code::InternalError, "no generated formula"))?;
                set_formulas(rule, vec![formula]);
            }
        },
        CfKind::ContainsBlanks
        | CfKind::NotContainsBlanks
        | CfKind::ContainsErrors
        | CfKind::NotContainsErrors => {
            if fresh || moved {
                let formula = generated_formula(kind, "", &first_cell(ranges))
                    .ok_or_else(|| error(Code::InternalError, "no generated formula"))?;
                set_formulas(rule, vec![formula]);
            }
        },
        CfKind::TimePeriod => {
            if let Some(template) = request.period {
                rule.template = template;
            }
            rule.parameter = rule
                .template
                .checked_sub(15)
                .and_then(|index| features::PERIODS.get(index as usize))
                .map_or(0, |(_, parameter)| *parameter);
            if fresh || request.period.is_some() || moved {
                set_formulas(
                    rule,
                    vec![period_formula(rule.template, &first_cell(ranges))],
                );
            }
        },
        CfKind::DataBar => {
            let mut bar = match rule.data_bar.take() {
                Some(bar) => bar,
                None => xlsb::conditional_formatting::Bar::new(
                    cf_threshold("min", None, 2)?,
                    cf_threshold("max", None, 3)?,
                    argb(0x0063_8EC6),
                ),
            };
            if request.min.is_some() {
                bar.min_cfvo = cf_threshold("min", request.min.as_deref(), 2)?;
            }
            if request.max.is_some() {
                bar.max_cfvo = cf_threshold("max", request.max.as_deref(), 3)?;
            }
            if let Some(color) = request.color {
                bar.color = argb(color);
                bar.color_record = xlsb::conditional_formatting::Color::from_argb(argb(color));
            }
            if let Some(value) = request.min_length {
                bar.min_length = value;
            }
            if let Some(value) = request.max_length {
                bar.max_length = value;
            }
            if let Some(value) = request.show_value {
                bar.show_value = value;
            }
            rule.data_bar = Some(bar);
        },
        CfKind::ColorScale => {
            let mut scale = match rule.color_scale.take() {
                Some(scale) => scale,
                None => xlsb::conditional_formatting::Scale::new(
                    cf_threshold("min", None, 2)?,
                    cf_threshold("max", None, 3)?,
                    argb(0x00F8_696B),
                    argb(0x0063_BE7B),
                ),
            };
            if (request.mid_color.is_some() || request.midpoint.is_some())
                && scale.mid_cfvo.is_none()
            {
                scale = scale.with_middle(
                    xlsb::conditional_formatting::Value::new(5, Some("50".to_owned())),
                    argb(0x00FF_EB84),
                );
            }
            if let Some(midpoint) = &request.midpoint {
                let mut threshold = cf_threshold("midpoint", Some(midpoint), 5)?;
                if threshold.cfvo_type == 1 && !midpoint.contains(':') {
                    threshold.cfvo_type = 5;
                }
                scale.mid_cfvo = Some(threshold);
            }
            if let Some(color) = request.mid_color {
                scale.mid_color = Some(argb(color));
                scale.mid_color_record =
                    Some(xlsb::conditional_formatting::Color::from_argb(argb(color)));
            }
            if let Some(color) = request.min_color {
                scale.min_color = argb(color);
                scale.min_color_record =
                    xlsb::conditional_formatting::Color::from_argb(argb(color));
            }
            if let Some(color) = request.max_color {
                scale.max_color = argb(color);
                scale.max_color_record =
                    xlsb::conditional_formatting::Color::from_argb(argb(color));
            }
            rule.color_scale = Some(scale);
        },
        CfKind::IconSet => {
            let mut icons = match rule.icon_set.take() {
                Some(icons) => icons,
                None => xlsb::conditional_formatting::IconSet::new(3, Vec::new()),
            };
            if let Some(set) = request.iconset {
                icons.icon_set_type = set;
                icons.cfvos.clear();
            }
            if icons.cfvos.is_empty() {
                let count: u32 = match icons.icon_set_type {
                    0..=7 => 3,
                    8..=12 => 4,
                    _ => 5,
                };
                icons.cfvos = (0..count)
                    .map(|step| {
                        xlsb::conditional_formatting::Value::new(
                            4,
                            Some((step * 100 / count).to_string()),
                        )
                    })
                    .collect();
            }
            if let Some(value) = request.reverse {
                icons.reverse = value;
            }
            if let Some(value) = request.show_value {
                icons.show_value = value;
            }
            rule.icon_set = Some(icons);
        },
        CfKind::DuplicateValues | CfKind::UniqueValues => {},
    }
    Ok(())
}

fn merge_style(mut base: DifferentialStyle, change: &DifferentialStyle) -> DifferentialStyle {
    if change.fill.is_some() {
        base.fill = change.fill;
    }
    if change.font_color.is_some() {
        base.font_color = change.font_color;
    }
    if change.bold.is_some() {
        base.bold = change.bold;
    }
    if change.italic.is_some() {
        base.italic = change.italic;
    }
    if change.strike.is_some() {
        base.strike = change.strike;
    }
    if change.underline.is_some() {
        base.underline = change.underline;
    }
    if change.size.is_some() {
        base.size = change.size;
    }
    if change.name.is_some() {
        base.name.clone_from(&change.name);
    }
    base
}

/// Whether a rule carries a differential format (visual rules do not).
const fn takes_style(rule: &Rule) -> bool {
    !matches!(
        rule.rule_type,
        RuleType::ColorScale | RuleType::DataBar | RuleType::IconSet
    )
}

impl Xlsb {
    fn add_cf(
        &self,
        parent: &Parent,
        implied: Option<&'static str>,
        props: &Props<'_>,
        max_snapshot_bytes: usize,
    ) -> Result<(Commit, Path<'static>)> {
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
            None => cf_ranges("ref", parent.reference.as_deref().unwrap_or("A1:A10"))?,
        };
        let editor = self.editor_for_write()?;
        let mut edit = editor
            .edit_conditional_formats_with_options(feature_edit_options!(
                conditional_formatting,
                max_snapshot_bytes
            ))
            .map_err(native)?;
        let worksheet_formats = editor
            .workbook()
            .map_err(native)?
            .sheet_features(parent.sheet, xlsb::sheet::FeatureSet::FORMATS)
            .map_err(native)?
            .formats;
        let priority = worksheet_formats
            .iter()
            .flat_map(|format| format.rules.iter().map(|rule| rule.priority))
            .max()
            .unwrap_or(0)
            .checked_add(1)
            .ok_or_else(|| error(Code::OutOfRange, "the priorities are exhausted"))?;
        let mut rule = fresh_rule(kind.kind, priority);
        if kind.percent {
            rule.percent = true;
        }
        if kind.bottom {
            rule.bottom = true;
        }
        if kind.below {
            rule.template = 26;
            rule.above_average = false;
        }
        apply_cf(&mut rule, &request, &ranges, true)?;
        if !style_is_empty(&request.style) {
            if !takes_style(&rule) {
                return Err(invalid(
                    "fill",
                    "a color scale, data bar or icon set takes no formatting",
                ));
            }
            rule.dxf_id = Some(edit.add_differential(&request.style).map_err(authored)?);
        }
        let mut format = Formatting::new(ranges);
        format.add_rule(rule);
        let index = edit.add(parent.sheet, format).map_err(authored)?;
        let commit = self.publish(edit.commit().map_err(native)?, 1, max_snapshot_bytes)?;
        Ok((commit, feature_path(parent.sheet, "cf", Some(index))))
    }

    fn set_cf(
        &self,
        sheet: usize,
        feature: &Feature,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        let Data::Cf(cf) = &feature.data else {
            return Err(error(
                Code::InternalError,
                "a conditional-format node holds no format",
            ));
        };
        let request = cf_request(Operation::Set, props)?;
        let mut format = cf.formatting.clone();
        if let Some(ranges) = &request.ranges {
            format.ranges.clone_from(ranges);
        }
        let ranges = format.ranges.clone();
        let rule = format.rules.first_mut().ok_or_else(|| {
            error(
                Code::UnsupportedCapability,
                "the conditional format has no rule",
            )
        })?;
        apply_cf(rule, &request, &ranges, false)?;
        // Formulas that only moved with their range are recompiled from the
        // text they read as, relative to the new first cell.
        if request.ranges.is_some() {
            for rule in format.rules.iter_mut().skip(1) {
                if !rule.formula_texts.is_empty() {
                    let texts = rule.formula_texts.clone();
                    set_formulas(rule, texts);
                }
            }
        }
        let editor = self.editor_for_write()?;
        let mut edit = editor
            .edit_conditional_formats_with_options(feature_edit_options!(
                conditional_formatting,
                limits.max_snapshot_bytes()
            ))
            .map_err(native)?;
        if !style_is_empty(&request.style) {
            let first = format
                .rules
                .first()
                .ok_or_else(|| error(Code::InternalError, "the rule vanished"))?;
            if !takes_style(first) {
                return Err(invalid(
                    "fill",
                    "a color scale, data bar or icon set takes no formatting",
                ));
            }
            let current = cf.dxf.clone().unwrap_or_default();
            let merged = merge_style(current, &request.style);
            let dxf = edit.add_differential(&merged).map_err(authored)?;
            if let Some(rule) = format.rules.first_mut() {
                rule.dxf_id = Some(dxf);
            }
        }
        edit.set(sheet, feature.index as usize, format)
            .map_err(authored)?;
        publish!(self, edit, limits)
    }
}

// ---------------------------------------------------------------------------
// Data validations

/// The rectangles of a validation's space-separated ranges.
fn rects(cell_ranges: &str) -> Vec<Area> {
    cell_ranges
        .split_whitespace()
        .filter_map(parse_area)
        .collect()
}

fn overlaps(left: &[Area], right: &[Area]) -> bool {
    left.iter().any(|a| {
        right
            .iter()
            .any(|b| a.0.0 <= b.1.0 && b.0.0 <= a.1.0 && a.0.1 <= b.1.1 && b.0.1 <= a.1.1)
    })
}

/// A validation formula: a leading `=` is dropped, and a list's literal
/// options are quoted.
fn validation_formula(key: &'static str, value: &str, list: bool) -> Result<String> {
    let value = value.trim();
    if value.len() > 255 {
        return Err(invalid(
            key,
            "a validation formula holds at most 255 characters",
        ));
    }
    Ok(if let Some(rest) = value.strip_prefix('=') {
        rest.to_owned()
    } else if list && !value.starts_with('"') {
        if value.contains('"') {
            return Err(invalid(key, "list options cannot contain quotes"));
        }
        quoted(value)
    } else {
        value.to_owned()
    })
}

fn validation_ranges(key: &'static str, value: &str) -> Result<String> {
    let ranges = value
        .split_whitespace()
        .map(|range| {
            parse_area(range)
                .map(area_label)
                .ok_or_else(|| invalid(key, "the range is not an A1 range"))
        })
        .collect::<Result<Vec<_>>>()?;
    if ranges.is_empty() {
        return Err(invalid(key, "the range is empty"));
    }
    Ok(ranges.join(" "))
}

fn apply_validation(
    validation: &mut xlsb::package::data_validation::Validation,
    props: &Props<'_>,
    operation: Operation,
) -> Result<()> {
    let admitted = admit(&features::VALIDATION_TABLE, operation, props)?;
    // The type first: a list's formula is quoted by type.
    for (key, name, arg) in &admitted {
        if *key == ValidationKey::Type {
            let value = arg.text().unwrap_or_default();
            let value = match value.to_ascii_lowercase().as_str() {
                "any" => "none".to_owned(),
                other => other.to_owned(),
            };
            let kind = features::VALIDATION_TYPES
                .iter()
                .position(|known| known.eq_ignore_ascii_case(&value))
                .and_then(|index| u8::try_from(index).ok())
                .ok_or_else(|| invalid(name, "the validation type is not one Excel knows"))?;
            validation.validation_type = kind;
        }
    }
    let list = validation.validation_type == 3;
    let mut formula1_set = false;
    for (key, name, arg) in admitted {
        let text = || arg.text().unwrap_or_default();
        match key {
            ValidationKey::Type => {},
            ValidationKey::Ref => validation.cell_ranges = validation_ranges(name, &text())?,
            ValidationKey::Operator => {
                let value = text();
                validation.operator = features::VALIDATION_OPERATORS
                    .iter()
                    .position(|known| known.eq_ignore_ascii_case(value.trim()))
                    .and_then(|index| u8::try_from(index).ok())
                    .ok_or_else(|| invalid(name, "the operator is not a validation comparison"))?;
            },
            ValidationKey::Formula1 => {
                formula1_set = true;
                validation.formula1 = arg
                    .text()
                    .map(|value| validation_formula(name, &value, list))
                    .transpose()?;
                validation.formula1_binary = None;
                validation.list_formula = None;
            },
            ValidationKey::Formula2 => {
                validation.formula2 = arg
                    .text()
                    .map(|value| validation_formula(name, &value, false))
                    .transpose()?;
                validation.formula2_binary = None;
            },
            ValidationKey::AllowBlank => validation.allow_blank = arg.flag(),
            ValidationKey::ShowError => validation.show_error_message = arg.flag(),
            ValidationKey::ShowInput => validation.show_input_message = arg.flag(),
            ValidationKey::ErrorTitle => validation.error_title = arg.text(),
            ValidationKey::Error => validation.error_text = arg.text(),
            ValidationKey::PromptTitle => validation.input_title = arg.text(),
            ValidationKey::Prompt => validation.input_text = arg.text(),
            ValidationKey::ErrorStyle => {
                let value = text().to_ascii_lowercase();
                let value = match value.as_str() {
                    "warn" => "warning",
                    "info" => "information",
                    other => other,
                };
                validation.error_style = features::ERROR_STYLES
                    .iter()
                    .position(|known| known.eq_ignore_ascii_case(value))
                    .and_then(|index| u8::try_from(index).ok())
                    .ok_or_else(|| invalid(name, "errorStyle is stop, warning or information"))?;
            },
            ValidationKey::InCellDropdown => validation.show_dropdown = arg.flag(),
        }
    }
    // A list of literal options is a quoted string list; a list naming
    // cells is an ordinary formula. A changed type recompiles the first
    // formula with the new kind's encoding.
    let string_list = list
        && validation
            .formula1
            .as_deref()
            .is_some_and(|formula| formula.starts_with('"'));
    if (formula1_set || validation.formula1_binary.is_some())
        && string_list != validation.string_list
    {
        validation.formula1_binary = None;
    }
    validation.string_list = string_list;
    for (formula, key) in [
        (&validation.formula1, "formula1"),
        (&validation.formula2, "formula2"),
    ] {
        if formula.as_deref().is_some_and(str::is_empty) {
            return Err(invalid(key, "a validation formula cannot be empty"));
        }
    }
    Ok(())
}

impl Xlsb {
    fn add_validation(
        &self,
        parent: &Parent,
        props: &Props<'_>,
        max_snapshot_bytes: usize,
    ) -> Result<(Commit, Path<'static>)> {
        let reference = props.iter().any(|(key, _)| {
            ["ref", "sqref", "range"]
                .iter()
                .any(|name| key.as_str().eq_ignore_ascii_case(name))
        });
        let fallback = match (&parent.reference, reference) {
            (_, true) => "A1".to_owned(),
            (Some(reference), false) => reference.clone(),
            (None, false) => {
                return Err(missing("ref", "a data validation needs a ref"));
            },
        };
        let mut validation = xlsb::package::data_validation::Validation::new(
            0,
            validation_ranges("ref", &fallback)?,
        );
        validation.allow_blank = true;
        validation.show_error_message = true;
        validation.show_input_message = true;
        apply_validation(&mut validation, props, Operation::Add)?;
        let editor = self.editor_for_write()?;
        let mut edit = editor
            .edit_data_validations_with_options(feature_edit_options!(
                data_validation,
                max_snapshot_bytes
            ))
            .map_err(native)?;
        let existing = edit.validations(parent.sheet).map_err(native)?;
        let added = rects(&validation.cell_ranges);
        if existing
            .iter()
            .any(|held| overlaps(&rects(&held.cell_ranges), &added))
        {
            return Err(error(
                Code::InvalidArgument,
                "the range overlaps an existing data validation",
            )
            .with_property("ref"));
        }
        let index = edit.add(parent.sheet, validation).map_err(authored)?;
        let commit = self.publish(edit.commit().map_err(native)?, 1, max_snapshot_bytes)?;
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
    ) -> Result<Commit> {
        let Data::Validation(current) = &feature.data else {
            return Err(error(
                Code::InternalError,
                "a validation node holds no validation",
            ));
        };
        let mut validation = (**current).clone();
        apply_validation(&mut validation, props, Operation::Set)?;
        let editor = self.editor_for_write()?;
        let mut edit = editor
            .edit_data_validations_with_options(feature_edit_options!(
                data_validation,
                limits.max_snapshot_bytes()
            ))
            .map_err(native)?;
        let existing = edit.validations(sheet).map_err(native)?;
        let changed = rects(&validation.cell_ranges);
        let clash = existing.iter().enumerate().any(|(index, other)| {
            index != feature.index as usize && overlaps(&rects(&other.cell_ranges), &changed)
        });
        if clash {
            return Err(error(
                Code::InvalidArgument,
                "the range overlaps another data validation",
            )
            .with_property("ref"));
        }
        edit.set(sheet, feature.index as usize, validation)
            .map_err(authored)?;
        publish!(self, edit, limits)
    }
}

// ---------------------------------------------------------------------------
// The auto-filter

fn filter_range(key: &'static str, value: &str) -> Result<xlsb::auto_filter::AutoFilter> {
    let ((first_row, first_column), (last_row, last_column)) =
        parse_area(value).ok_or_else(|| invalid(key, "the range is not an A1 range"))?;
    Ok(xlsb::auto_filter::AutoFilter {
        first_row,
        last_row,
        first_column,
        last_column,
    })
}

impl Xlsb {
    fn add_autofilter(
        &self,
        parent: &Parent,
        props: &Props<'_>,
        max_snapshot_bytes: usize,
    ) -> Result<(Commit, Path<'static>)> {
        let mut range = parent.reference.clone();
        for (key, _, arg) in admit(&features::AUTOFILTER_TABLE, Operation::Add, props)? {
            if key == AutofilterKey::Range {
                range = arg.text();
            }
        }
        let range = range.ok_or_else(|| missing("range", "an auto-filter needs a range"))?;
        let filter = filter_range("range", &range)?;
        let editor = self.editor_for_write()?;
        let mut edit = editor
            .edit_auto_filter_with_options(feature_edit_options!(auto_filter, max_snapshot_bytes))
            .map_err(native)?;
        edit.set_auto_filter(parent.sheet, Some(filter))
            .map_err(authored)?;
        let commit = self.publish(edit.commit().map_err(native)?, 1, max_snapshot_bytes)?;
        Ok((commit, feature_path(parent.sheet, "autofilter", None)))
    }

    fn set_autofilter(
        &self,
        sheet: usize,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        let editor = self.editor_for_write()?;
        let mut edit = editor
            .edit_auto_filter_with_options(feature_edit_options!(
                auto_filter,
                limits.max_snapshot_bytes()
            ))
            .map_err(native)?;
        let mut filter = edit
            .auto_filter(sheet)
            .map_err(native)?
            .ok_or_else(|| error(Code::PathNotFound, "the worksheet has no auto-filter"))?;
        for (key, name, arg) in admit(&features::AUTOFILTER_TABLE, Operation::Set, props)? {
            if key == AutofilterKey::Range {
                filter = filter_range(name, &arg.text().unwrap_or_default())?;
            }
        }
        edit.set_auto_filter(sheet, Some(filter))
            .map_err(authored)?;
        publish!(self, edit, limits)
    }
}

// ---------------------------------------------------------------------------
// Page breaks

fn break_index(kind: Kind, key: &'static str, arg: &Arg) -> Result<u32> {
    let (maximum, letters) = if kind == Kind::RowBreak {
        (i64::from(xlsb::page_breaks::LAST_ROW), false)
    } else {
        (i64::from(xlsb::page_breaks::LAST_COLUMN), true)
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
    u32::try_from(value)
        .ok()
        .filter(|value| (1..=maximum).contains(&i64::from(*value)))
        .ok_or_else(|| invalid(key, "the break is out of the worksheet"))
}

impl Xlsb {
    fn breaks(&self, sheet: usize, kind: Kind) -> Result<Vec<xlsb::page_breaks::Break>> {
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
        breaks: Vec<xlsb::page_breaks::Break>,
        max_snapshot_bytes: usize,
    ) -> Result<Commit> {
        let editor = self.editor_for_write()?;
        let mut edit = editor
            .edit_page_breaks_with_options(feature_edit_options!(page_breaks, max_snapshot_bytes))
            .map_err(native)?;
        let mut all = edit.page_breaks(sheet).map_err(native)?;
        if kind == Kind::RowBreak {
            all.rows = breaks;
        } else {
            all.columns = breaks;
        }
        edit.set_page_breaks(sheet, &all).map_err(authored)?;
        let commit = edit.commit().map_err(native)?;
        self.publish(commit, 1, max_snapshot_bytes)
    }

    fn break_spec(
        kind: Kind,
        current: Option<xlsb::page_breaks::Break>,
        props: &Props<'_>,
        operation: Operation,
    ) -> Result<xlsb::page_breaks::Break> {
        let full = if kind == Kind::RowBreak {
            xlsb::page_breaks::LAST_COLUMN
        } else {
            xlsb::page_breaks::LAST_ROW
        };
        let mut id = current.map(|item| item.id);
        let mut minimum = current.map_or(0, |item| item.min);
        let mut maximum = current.map_or(full, |item| item.max);
        let mut manual = current.is_none_or(|item| item.manual);
        let admitted = if kind == Kind::RowBreak {
            admit(&features::ROW_BREAK_TABLE, operation, props)?
        } else {
            admit(&features::COL_BREAK_TABLE, operation, props)?
        };
        for (key, name, arg) in admitted {
            match key {
                BreakKey::Index => id = Some(break_index(kind, name, &arg)?),
                BreakKey::Manual => manual = arg.flag(),
                BreakKey::Min | BreakKey::Max => {
                    let value = u32::try_from(arg.int(name)?)
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
        let id = id.ok_or_else(|| missing(key, "a page break needs its row or column"))?;
        if !manual {
            return Err(invalid(
                "manual",
                "an XLSB worksheet stores manual page breaks only",
            ));
        }
        if minimum > maximum || maximum > full {
            return Err(invalid("max", "the break span is outside the worksheet"));
        }
        Ok(xlsb::page_breaks::Break {
            id,
            min: minimum,
            max: maximum,
            manual,
            pivot: current.is_some_and(|item| item.pivot),
        })
    }

    fn add_break(
        &self,
        parent: &Parent,
        kind: Kind,
        props: &Props<'_>,
        max_snapshot_bytes: usize,
    ) -> Result<(Commit, Path<'static>)> {
        let item = Self::break_spec(kind, None, props, Operation::Add)?;
        let mut breaks = self.breaks(parent.sheet, kind)?;
        if breaks.iter().any(|held| held.id == item.id) {
            return Err(error(
                Code::InvalidArgument,
                "the worksheet already has this page break",
            ));
        }
        let index = breaks.partition_point(|held| held.id < item.id);
        breaks.insert(index, item);
        let commit = self.publish_breaks(parent.sheet, kind, breaks, max_snapshot_bytes)?;
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
    ) -> Result<Commit> {
        let Data::Break(current) = feature.data else {
            return Err(error(
                Code::InternalError,
                "a page-break node holds no break",
            ));
        };
        let item = Self::break_spec(feature.kind, Some(current), props, Operation::Set)?;
        let mut breaks = self.breaks(sheet, feature.kind)?;
        remove_at(&mut breaks, feature.index)?;
        if breaks.iter().any(|held| held.id == item.id) {
            return Err(error(
                Code::InvalidArgument,
                "the worksheet already has this page break",
            ));
        }
        let index = breaks.partition_point(|held| held.id < item.id);
        breaks.insert(index, item);
        self.publish_breaks(sheet, feature.kind, breaks, limits.max_snapshot_bytes())
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
