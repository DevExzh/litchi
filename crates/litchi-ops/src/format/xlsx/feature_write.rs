//! Add, set and remove of worksheet features (record 0320).
//!
//! Every verb admits its properties through the kind's registry table in
//! `features`, builds the leaf value, and publishes one leaf commit:
//!
//! | Kind | Leaf transaction |
//! |---|---|
//! | `comment` | `workbook::comments::edit` (splices notes and VML note shapes) |
//! | `hyperlink`, cell `link`/`tooltip`/`display` | the primary `workbook::Edit` (`put_hyperlink`, `remove_hyperlink`) |
//! | `picture`, `shape` | `drawing::edit` |
//! | `cf` | `conditional_formatting::edit` (with a differential format for its formatting) |
//! | `dataValidation` | `data_validation::edit` |
//! | `autofilter` | `auto_filter::edit` |
//! | `rowbreak`, `colbreak` | `Package::edit_page_breaks` |
//!
//! Property names and defaults follow OfficeCLI's Excel handler; where this
//! adapter is stricter (an unknown CF `type`, a list literal with quotes) it
//! refuses rather than guesses.

use std::io::Read as _;

use litchi_core::selector::{Path, Predicate, Step};
use litchi_core::{FileFormat, Position};
use litchi_xlsx as xlsx;
use xlsx::conditional_formatting::edit::{DifferentialStyle, Underline};
use xlsx::conditional_formatting::{
    Color as CfColor, ColorRole, ColorScale, DataBar, DifferentialRef, Direction, Formatting,
    IconSet, Icons, Kind as RuleKind, NamedColor, Operator, Payload, Period, Range as CfRange,
    Rule, Value as Threshold, ValueKind,
};
use xlsx::drawing::edit::{
    Align, COLUMN_EMU, Change, ElementKind, Fill, Font, Line, Marker, PictureSpec, Placement,
    ROW_EMU, ShapeSpec, VerticalAlign,
};
use xlsx::workbook::comments::{RunColor, RunFont, TextRun};

use super::features::{
    self, AutofilterKey, BreakKey, CfKey, CommentKey, Data, Feature, HyperlinkKey, PictureKey,
    ShapeKey, ValidationKey,
};
use super::node::{Item, Node};
use super::spec::Kind;
use super::write::{Commit, edit_with_snapshot_limit};
use super::{Nav, Xlsx};
use crate::error::{Code, Error, Result};
use crate::new_node::NewNode;
use crate::placement::Placement as NodePlacement;
use crate::props::Props;
use crate::registry::{Operation, Table};
use crate::target::TargetArg;
use crate::transaction::SemanticLimits;
use crate::value::{Color, ColorContext, Length, Scalar};

/// The package and semantic patch charge returned by one native feature
/// family. These commits all use the shared part-level patch engine, but keep
/// their concrete commit types in their per-family modules.
trait NativeFeatureCommit {
    fn into_feature_parts(self) -> (xlsx::Package, bool, usize);
}

macro_rules! native_feature_commit {
    ($commit:ty) => {
        impl NativeFeatureCommit for $commit {
            fn into_feature_parts(self) -> (xlsx::Package, bool, usize) {
                let changed = self.diagnostics().changed();
                let patch_bytes = self.patch().retained_bytes();
                (self.into_package(), changed, patch_bytes)
            }
        }
    };
}

native_feature_commit!(xlsx::workbook::comments::edit::Commit);
native_feature_commit!(xlsx::drawing::edit::Commit);
native_feature_commit!(xlsx::conditional_formatting::edit::Commit);
native_feature_commit!(xlsx::data_validation::edit::Commit);
native_feature_commit!(xlsx::auto_filter::edit::Commit);
native_feature_commit!(xlsx::page_breaks::edit::Commit);

/// Reserve both package endpoints before admitting a part-family patch. The
/// patch editor then checks its own exact semantic payload before constructing
/// patch-owned states; the final adapter repeats the complete charge.
pub(super) fn part_edit_options(source: &xlsx::Package, maximum: usize) -> xlsx::EditOptions {
    let _ = source;
    xlsx::EditOptions::default().with_max_retained_bytes(maximum)
}

fn error(code: Code, message: impl Into<String>) -> Error {
    Error::new(code, message.into()).with_format(FileFormat::Xlsx)
}

fn invalid(key: &'static str, message: impl Into<String>) -> Error {
    error(Code::InvalidValue, message).with_property(key)
}

/// A leaf refusal of authored input: an invalid value, not a corrupt file.
fn authored(error: xlsx::Error) -> Error {
    match error {
        xlsx::Error::Invalid(message) => {
            Error::new(Code::InvalidValue, message).with_format(FileFormat::Xlsx)
        },
        error => Error::from(error),
    }
}

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
        Item::Cell(_) => {
            let (_, at) = node.write_cell_target().ok_or_else(unsupported_parent)?;
            Ok(Parent {
                sheet,
                reference: Some(at.a1()),
            })
        },
        Item::Range(_) => {
            let (_, rect) = node.write_range_target().ok_or_else(unsupported_parent)?;
            Ok(Parent {
                sheet,
                reference: Some(rect.a1()),
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

impl Xlsx {
    /// Publish a feature probe under the transaction's remaining native
    /// snapshot budget. Add routes retain the legacy unconstrained adapter,
    /// while set routes must check the staged and canonical signed endpoints
    /// before the private probe is discarded.
    fn publish_with_limit(
        &self,
        source: xlsx::Package,
        target: xlsx::Package,
        changed: bool,
        native_patch_bytes: usize,
        max_snapshot_bytes: usize,
    ) -> Result<Commit> {
        if changed && self.signature() == super::Signature::Signed {
            return Err(Error::signed(FileFormat::Xlsx));
        }
        let after = if changed {
            self.adopt_workbook(target.clone())
        } else {
            self.clone()
        };
        Commit::from_metadata(
            self,
            after,
            source,
            target,
            changed,
            1,
            native_patch_bytes,
            max_snapshot_bytes,
        )
    }

    /// The features of `kind` on the worksheet at `sheet`, read again.
    fn features_of(&self, sheet: usize, kind: Kind) -> Result<Vec<Feature>> {
        let handle = super::node::SheetHandle::at(self, sheet)?
            .ok_or_else(|| error(Code::PathNotFound, "the worksheet is missing"))?;
        let sheet = super::node::Sheet::new(self, sheet, handle);
        features::list(self, &sheet, kind)
    }

    /// Add a worksheet feature under `parent`; `implied` is the rule type a
    /// CF element word (`databar`, `cellis`) names.
    pub(super) fn add_feature(
        &self,
        parent_node: &Node<'_>,
        kind: Kind,
        implied: Option<&'static str>,
        node: NewNode<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        if !matches!(node.placement(), None | Some(NodePlacement::End)) {
            return Err(error(
                Code::UnsupportedCapability,
                "worksheet features are added at the end of their kind",
            )
            .with_suggestion("omit placement or use Placement::End"));
        }
        let parent = parent(parent_node)?;
        match kind {
            Kind::Comment => self.add_comment(&parent, node.props(), limits),
            Kind::Hyperlink => self.add_hyperlink(&parent, node.props(), limits),
            Kind::Picture => self.add_picture(&parent, node.props(), limits),
            Kind::Shape => self.add_shape(&parent, node.props(), limits),
            Kind::Cf => self.add_cf(&parent, implied, node.props(), limits),
            Kind::DataValidation => self.add_validation(&parent, node.props(), limits),
            Kind::Autofilter => self.add_autofilter(&parent, node.props(), limits),
            Kind::RowBreak | Kind::ColBreak => self.add_break(&parent, kind, node.props(), limits),
            _ => Err(error(
                Code::UnsupportedElement,
                "this element kind cannot be added to an XLSX worksheet",
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
        let source = self.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        let (target, changed, native_patch_bytes) = match &feature.data {
            Data::Comment(note) => {
                let mut edit = source.edit_comments_with_options(part_edit_options(
                    &source,
                    limits.max_snapshot_bytes(),
                ))?;
                edit.remove(sheet, &note.reference)?;
                split(edit.commit()?)
            },
            Data::Hyperlink(link) => {
                let mut edit = edit_with_snapshot_limit(&source, limits.max_snapshot_bytes())?;
                edit.remove_hyperlink(sheet, link.reference().clone())?;
                let native = edit.commit()?;
                return Ok(Some(Commit::from_native_with_limit(
                    self,
                    source,
                    native,
                    limits.max_snapshot_bytes(),
                )?));
            },
            Data::Drawing(_) => {
                let kind = if feature.kind == Kind::Picture {
                    ElementKind::Picture
                } else {
                    ElementKind::Shape
                };
                let mut edit = source.edit_drawings_with_options(part_edit_options(
                    &source,
                    limits.max_snapshot_bytes(),
                ))?;
                edit.remove(sheet, kind, feature.index as usize)?;
                let commit = edit.commit()?;
                let changed = commit.diagnostics().changed();
                let native_patch_bytes = commit.patch().retained_bytes();
                (commit.into_package(), changed, native_patch_bytes)
            },
            Data::Cf(_) => {
                let mut edit = source.edit_conditional_formats_with_options(part_edit_options(
                    &source,
                    limits.max_snapshot_bytes(),
                ))?;
                let mut formats = edit.formats(sheet)?;
                remove_at(&mut formats, feature.index)?;
                edit.set_formats(sheet, formats).map_err(authored)?;
                let commit = edit.commit()?;
                let changed = commit.diagnostics().changed();
                let native_patch_bytes = commit.patch().retained_bytes();
                (commit.into_package(), changed, native_patch_bytes)
            },
            Data::Validation(_) => {
                let mut edit = source.edit_data_validations_with_options(part_edit_options(
                    &source,
                    limits.max_snapshot_bytes(),
                ))?;
                let mut collections = edit.validations(sheet)?;
                let (collection, item) = locate(&collections, feature.index)?;
                let mut validations = collections[collection].validations().to_vec();
                validations.remove(item);
                if validations.is_empty() {
                    collections.remove(collection);
                } else {
                    collections[collection]
                        .set_validations(validations)
                        .map_err(authored)?;
                }
                edit.set_validations(sheet, collections).map_err(authored)?;
                let commit = edit.commit()?;
                let changed = commit.diagnostics().changed();
                let native_patch_bytes = commit.patch().retained_bytes();
                (commit.into_package(), changed, native_patch_bytes)
            },
            Data::Autofilter(_) => {
                let mut edit = source.edit_auto_filter_with_options(part_edit_options(
                    &source,
                    limits.max_snapshot_bytes(),
                ))?;
                edit.set_auto_filter(sheet, None)?;
                let commit = edit.commit()?;
                let changed = commit.diagnostics().changed();
                let native_patch_bytes = commit.patch().retained_bytes();
                (commit.into_package(), changed, native_patch_bytes)
            },
            Data::Break(_) => {
                let mut breaks = self.breaks(sheet, feature.kind)?;
                remove_at(&mut breaks, feature.index)?;
                return self
                    .publish_breaks(
                        source,
                        sheet,
                        feature.kind,
                        breaks,
                        limits.max_snapshot_bytes(),
                    )
                    .map(Some);
            },
        };
        Ok(Some(self.publish_with_limit(
            source,
            target,
            changed,
            native_patch_bytes,
            limits.max_snapshot_bytes(),
        )?))
    }
}

fn split<C: NativeFeatureCommit>(commit: C) -> (xlsx::Package, bool, usize) {
    commit.into_feature_parts()
}

fn remove_at<T>(items: &mut Vec<T>, index: u32) -> Result<T> {
    let index = index as usize;
    if index >= items.len() {
        return Err(error(Code::PathNotFound, "the feature is missing"));
    }
    Ok(items.remove(index))
}

// ---------------------------------------------------------------------------
// Notes

fn check_cell(key: &'static str, reference: &str) -> Result<String> {
    let cell = litchi_core::selector::CellRef::parse(reference.trim().trim_start_matches('$'))
        .ok_or_else(|| invalid(key, "the reference is not one A1 cell"))?;
    Ok(super::node::Label::cell(cell.column(), cell.row())
        .as_str()
        .to_owned())
}

/// The `font.*` keys of one note write, which become the run properties
/// of the single run OfficeCLI writes for them.
#[derive(Default)]
struct NoteFont {
    given: bool,
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<String>,
    size: Option<f64>,
    color: Option<xlsx::Rgb>,
    name: Option<String>,
}

impl NoteFont {
    /// Admit one `font.*` key; `false` for any other key.
    fn admit(&mut self, key: CommentKey, name: &'static str, arg: &Arg) -> Result<bool> {
        match key {
            CommentKey::FontBold => self.bold = Some(arg.flag()),
            CommentKey::FontItalic => self.italic = Some(arg.flag()),
            CommentKey::FontUnderline => self.underline = note_underline(name, arg)?,
            CommentKey::FontSize => self.size = Some(note_size(name, arg)?),
            CommentKey::FontColor => {
                self.color = Some(arg_color(name, arg, ColorContext::Spreadsheet)?);
            },
            CommentKey::FontName => self.name = arg.text(),
            _ => return Ok(false),
        }
        self.given = true;
        Ok(true)
    }

    /// The one run of `text` OfficeCLI writes for `font.*` keys: Excel's
    /// note size and typeface, and its note color (`indexed` 81), unless
    /// the keys name others.
    fn run(self, text: String) -> TextRun {
        TextRun {
            text,
            font: RunFont {
                bold: self.bold.filter(|on| *on),
                italic: self.italic.filter(|on| *on),
                underline: self.underline,
                size: Some(self.size.unwrap_or(features::NOTE_SIZE)),
                color: Some(self.color.map_or(RunColor::Indexed(81), argb)),
                name: Some(self.name.unwrap_or_else(|| features::NOTE_FONT.to_owned())),
                ..RunFont::default()
            },
        }
    }
}

fn argb(color: xlsx::Rgb) -> RunColor {
    RunColor::Rgb(format!(
        "FF{:02X}{:02X}{:02X}",
        color.red(),
        color.green(),
        color.blue()
    ))
}

/// OfficeCLI's note underline: `double`, or `single` for any other value
/// except `none`/`false`.
fn note_underline(key: &'static str, arg: &Arg) -> Result<Option<String>> {
    let value = match arg {
        Arg::Bool(on) => return Ok(on.then(|| "single".to_owned())),
        other => other.text().unwrap_or_default(),
    };
    Ok(match value.trim().to_ascii_lowercase().as_str() {
        "" | "none" | "false" => None,
        "double" => Some("double".to_owned()),
        "single" | "true" => Some("single".to_owned()),
        _ => return Err(invalid(key, "underline is single, double or none")),
    })
}

fn note_size(key: &'static str, arg: &Arg) -> Result<f64> {
    match arg {
        Arg::Length(length) => Ok(length.points()),
        other => other.number(key),
    }
}

/// Parse OfficeCLI's `runs` JSON: an array of run objects with `text`,
/// `bold`, `italic`, `strike`, `superscript`, `subscript`, `underline`,
/// `size`, `color` and `font`; every run gets Excel's note size, color and
/// typeface unless it names others.
fn note_runs(key: &'static str, value: &str) -> Result<Vec<TextRun>> {
    let parsed: serde_json::Value =
        serde_json::from_str(value).map_err(|_| invalid(key, "runs is not JSON"))?;
    let serde_json::Value::Array(items) = parsed else {
        return Err(invalid(key, "runs is a JSON array of run objects"));
    };
    let truthy = |value: Option<&serde_json::Value>| match value {
        Some(serde_json::Value::Bool(on)) => *on,
        Some(serde_json::Value::String(text)) => {
            matches!(text.to_ascii_lowercase().as_str(), "true" | "1" | "yes")
        },
        _ => false,
    };
    let string = |value: Option<&serde_json::Value>| match value {
        Some(serde_json::Value::String(text)) => Some(text.clone()),
        Some(serde_json::Value::Number(number)) => Some(number.to_string()),
        _ => None,
    };
    let mut runs = Vec::new();
    for item in &items {
        let serde_json::Value::Object(item) = item else {
            continue;
        };
        let mut font = NoteFont::default();
        if let Some(size) = string(item.get("size")) {
            let size = size.trim().trim_end_matches("pt").trim();
            font.size = Some(
                size.parse()
                    .map_err(|_| invalid(key, "a run size is a number of points"))?,
            );
        }
        if let Some(color) = string(item.get("color")) {
            font.color = Some(parse_color(key, &color, ColorContext::Spreadsheet)?);
        }
        if let Some(underline) = string(item.get("underline")) {
            font.underline = note_underline(key, &Arg::Text(underline))?;
        }
        font.name = string(item.get("font"));
        font.bold = Some(truthy(item.get("bold")));
        font.italic = Some(truthy(item.get("italic")));
        let mut run = font.run(string(item.get("text")).unwrap_or_default());
        run.font.strike = truthy(item.get("strike")).then_some(true);
        run.font.vert_align = if truthy(item.get("superscript")) {
            Some("superscript".to_owned())
        } else if truthy(item.get("subscript")) {
            Some("subscript".to_owned())
        } else {
            None
        };
        runs.push(run);
    }
    if runs.is_empty() {
        return Err(invalid(key, "runs names no run object"));
    }
    Ok(runs)
}

impl Xlsx {
    fn add_comment(
        &self,
        parent: &Parent,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        let mut reference = None;
        let mut author = None;
        let mut text = None;
        let mut font = NoteFont::default();
        let mut runs = None;
        for (key, name, arg) in admit(&features::COMMENT_TABLE, Operation::Add, props)? {
            if font.admit(key, name, &arg)? {
                continue;
            }
            match key {
                CommentKey::Ref => {
                    reference = Some(check_cell(name, &arg.text().unwrap_or_default())?)
                },
                CommentKey::Author => author = arg.text(),
                CommentKey::Text => text = arg.text(),
                // OfficeCLI: a non-blank `runs` wins over `text` and `font.*`.
                CommentKey::Runs => {
                    runs = arg
                        .text()
                        .filter(|value| !value.trim().is_empty())
                        .map(|value| note_runs(name, &value))
                        .transpose()?;
                },
                _ => {},
            }
        }
        let reference = match (reference, &parent.reference) {
            (Some(reference), _) => reference,
            (None, Some(cell)) => check_cell("ref", cell)?,
            (None, None) => {
                return Err(error(Code::MissingProperty, "a note needs a ref").with_property("ref"));
            },
        };
        let source = self.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        let mut edit = source
            .edit_comments_with_options(part_edit_options(&source, limits.max_snapshot_bytes()))?;
        let author = author.as_deref().unwrap_or("Author");
        let runs = match runs {
            Some(runs) => Some(runs),
            None if font.given => Some(vec![font.run(text.clone().unwrap_or_default())]),
            None => None,
        };
        let index = match &runs {
            Some(runs) => edit.add_runs(parent.sheet, &reference, author, runs),
            None => edit.add(
                parent.sheet,
                &reference,
                author,
                text.as_deref().unwrap_or(""),
            ),
        }
        .map_err(authored)?;
        // Notes are numbered row-major, as `features::list` numbers them.
        let Some(index) = index else {
            return Err(error(Code::InvalidArgument, "the cell already has a note")
                .with_property("ref")
                .with_suggestion("set the existing comment instead"));
        };
        let (target, changed, native_patch_bytes) = split(edit.commit()?);
        let commit = self.publish_with_limit(
            source,
            target,
            changed,
            native_patch_bytes,
            limits.max_snapshot_bytes(),
        )?;
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
        let mut reference = note.reference.clone();
        let mut author = note.author.clone();
        let mut text = note.text.clone();
        let mut font = NoteFont::default();
        let mut runs = None;
        for (key, name, arg) in admit(&features::COMMENT_TABLE, Operation::Set, props)? {
            if font.admit(key, name, &arg)? {
                continue;
            }
            match key {
                CommentKey::Ref => reference = check_cell(name, &arg.text().unwrap_or_default())?,
                CommentKey::Author => author = arg.text().unwrap_or_default(),
                CommentKey::Text => text = arg.text().unwrap_or_default(),
                CommentKey::Runs => {
                    runs = Some(note_runs(name, &arg.text().unwrap_or_default())?);
                },
                _ => {},
            }
        }
        // `font.*` rebuilds the note as one run formatted by this call's keys
        // alone, keeping its text (OfficeCLI); a new `text` alone writes plain
        // text; a rich note whose text stays keeps its runs.
        let runs = match runs {
            Some(runs) => Some(runs),
            None if font.given => Some(vec![font.run(text.clone())]),
            None if text == note.text && !note.runs.is_empty() => Some(note.runs.clone()),
            None => None,
        };
        let source = self.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        let mut edit = source
            .edit_comments_with_options(part_edit_options(&source, limits.max_snapshot_bytes()))?;
        if reference != note.reference {
            if edit
                .comments(sheet)?
                .is_some_and(|comments| comments.get(&reference).is_some())
            {
                return Err(
                    error(Code::InvalidArgument, "the target cell already has a note")
                        .with_property("ref"),
                );
            }
            edit.remove(sheet, &note.reference)?;
        }
        match &runs {
            Some(runs) => edit.set_runs(sheet, &reference, &author, runs),
            None => edit.set(sheet, &reference, &author, &text),
        }
        .map_err(authored)?;
        let (target, changed, native_patch_bytes) = split(edit.commit()?);
        self.publish_with_limit(
            source,
            target,
            changed,
            native_patch_bytes,
            limits.max_snapshot_bytes(),
        )
    }
}

fn commit_file(commit: &Commit) -> &Xlsx {
    commit.file()
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

/// The hyperlink properties of one cell `set` (`link`, `tooltip`,
/// `display`): `Some(None)` removes a link.
#[derive(Default)]
pub(super) struct CellLink {
    link: Option<Option<String>>,
    tooltip: Option<Option<String>>,
    display: Option<Option<String>>,
}

impl CellLink {
    fn is_empty(&self) -> bool {
        self.link.is_none() && self.tooltip.is_none() && self.display.is_none()
    }
}

/// Split a cell `set`'s hyperlink properties (`spec::CELL_LINK_TABLE`)
/// from the rest.
pub(super) fn split_cell_link<'a>(props: &'a Props<'a>) -> Result<(Option<CellLink>, Props<'a>)> {
    let mut link = CellLink::default();
    let mut rest = Props::with_capacity(props.len());
    for (key, value) in props.iter() {
        let Some((slot, entry)) = super::spec::CELL_LINK_TABLE.lookup(key.as_str()) else {
            rest.push(key.reborrow(), value.reborrow());
            continue;
        };
        let name = entry.spec.name;
        let text = entry.prepare(Operation::Set, value, |_, scalar| {
            Ok(match scalar {
                Some(Scalar::Text(text)) => Some(text.to_owned()),
                _ => None,
            })
        })?;
        let text = text
            .filter(|text| !(slot == 0 && (text.is_empty() || text.eq_ignore_ascii_case("none"))));
        let field = match slot {
            0 => &mut link.link,
            1 => &mut link.tooltip,
            _ => &mut link.display,
        };
        if field.is_some() {
            return Err(error(
                Code::InvalidArgument,
                "a property was assigned more than once",
            )
            .with_property(name));
        }
        *field = Some(text);
    }
    Ok(((!link.is_empty()).then_some(link), rest))
}

impl Xlsx {
    fn hyperlink_position(&self, sheet: usize, reference: &str) -> Result<usize> {
        self.features_of(sheet, Kind::Hyperlink)?
            .iter()
            .position(|feature| {
                matches!(&feature.data, Data::Hyperlink(link) if link.reference().as_str().eq_ignore_ascii_case(reference))
            })
            .ok_or_else(|| error(Code::InternalError, "the added hyperlink did not read back"))
    }

    fn add_hyperlink(
        &self,
        parent: &Parent,
        props: &Props<'_>,
        limits: &SemanticLimits,
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
        let reference = xlsx::hyperlinks::HyperlinkReference::new(&reference)
            .map_err(|error| authored(error).with_property("ref"))?;
        let link =
            xlsx::hyperlinks::Hyperlink::new(reference.clone(), location, url, display, tooltip)
                .map_err(|error| authored(error).with_property("url"))?;
        let source = self.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        let mut edit = edit_with_snapshot_limit(&source, limits.max_snapshot_bytes())?;
        edit.put_hyperlink(parent.sheet, link).map_err(authored)?;
        let native = edit.commit()?;
        let commit =
            Commit::from_native_with_limit(self, source, native, limits.max_snapshot_bytes())?;
        let index = commit_file(&commit).hyperlink_position(parent.sheet, reference.as_str())?;
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
        let mut reference = link.reference().as_str().to_owned();
        let mut location = link.location().map(str::to_owned);
        let mut url = link.external_target().map(str::to_owned);
        let mut display = link.display().map(str::to_owned);
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
        let new_reference = xlsx::hyperlinks::HyperlinkReference::new(&reference)
            .map_err(|error| authored(error).with_property("ref"))?;
        let replacement =
            xlsx::hyperlinks::Hyperlink::new(new_reference, location, url, display, tooltip)
                .map_err(|error| authored(error).with_property("url"))?;
        let source = self.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        let mut edit = edit_with_snapshot_limit(&source, limits.max_snapshot_bytes())?;
        if replacement.reference() != link.reference() {
            edit.remove_hyperlink(sheet, link.reference().clone())?;
        }
        edit.put_hyperlink(sheet, replacement).map_err(authored)?;
        let native = edit.commit()?;
        Commit::from_native_with_limit(self, source, native, limits.max_snapshot_bytes())
    }

    /// Stage a cell `set`'s hyperlink properties into the cell edit.
    pub(super) fn stage_cell_link(
        &self,
        edit: &mut xlsx::workbook::Edit,
        sheet: usize,
        at: xlsx::Address,
        change: &CellLink,
    ) -> Result<()> {
        let reference = at.a1();
        let link_reference = xlsx::hyperlinks::HyperlinkReference::new(&reference)?;
        // As OfficeCLI: a new link replaces the cell's hyperlink, taking its
        // tooltip and display text from the same command only; removing one
        // needs no read either. Only a tooltip or display change alone reads
        // the existing hyperlink.
        let (location, url, mut display, mut tooltip) = match &change.link {
            Some(Some(value)) => {
                let (location, url) = link_target("link", value)?;
                (location, url, None, None)
            },
            Some(None) => {
                edit.remove_hyperlink(sheet, link_reference)?;
                return Ok(());
            },
            None => {
                let existing = super::node::SheetHandle::at(self, sheet)?
                    .ok_or_else(|| error(Code::PathNotFound, "the worksheet is missing"))?
                    .hyperlinks()?
                    .into_iter()
                    .find(|link| link.reference().as_str() == reference);
                existing.map_or((None, None, None, None), |link| {
                    (
                        link.location().map(str::to_owned),
                        link.external_target().map(str::to_owned),
                        link.display().map(str::to_owned),
                        link.tooltip().map(str::to_owned),
                    )
                })
            },
        };
        if location.is_none() && url.is_none() {
            return Err(error(
                Code::UnsupportedProperty,
                "a tooltip or display text needs a link on the cell",
            )
            .with_property(if change.tooltip.is_some() {
                "tooltip"
            } else {
                "display"
            })
            .with_suggestion("set link in the same command"));
        }
        if let Some(value) = &change.tooltip {
            tooltip.clone_from(value);
        }
        if let Some(value) = &change.display {
            display.clone_from(value);
        }
        let link =
            xlsx::hyperlinks::Hyperlink::new(link_reference, location, url, display, tooltip)
                .map_err(|error| authored(error).with_property("link"))?;
        edit.put_hyperlink(sheet, link)?;
        Ok(())
    }
}

// ---------------------------------------------------------------------------
// Pictures and shapes

/// A picture source: a base64 `data:` URI or a local file path.
fn read_source(src: &str, staged_limit: usize) -> Result<Vec<u8>> {
    const NATIVE_MAXIMUM: usize = xlsx::drawing::edit::MAX_IMAGE_BYTES;
    let maximum = staged_limit.min(NATIVE_MAXIMUM);
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
        let significant = payload
            .bytes()
            .filter(|byte| !byte.is_ascii_whitespace())
            .count();
        let padding = payload
            .bytes()
            .rev()
            .skip_while(u8::is_ascii_whitespace)
            .take_while(|byte| *byte == b'=')
            .count();
        let decoded = significant
            .saturating_add(3)
            .checked_div(4)
            .unwrap_or(usize::MAX)
            .saturating_mul(3)
            .saturating_sub(if significant % 4 == 0 {
                padding.min(2)
            } else {
                0
            });
        if decoded > maximum {
            return Err(invalid("src", "the image is too large"));
        }
        let mut decoder = base64::read::DecoderReader::new(
            Base64Input::new(payload.as_bytes()),
            &base64::engine::general_purpose::STANDARD,
        );
        let mut output = Vec::with_capacity(decoded.min(maximum));
        let mut chunk = [0_u8; 8192];
        loop {
            let count = decoder
                .read(&mut chunk)
                .map_err(|_| invalid("src", "the data URI payload is not base64"))?;
            if count == 0 {
                break;
            }
            let observed = output.len().saturating_add(count);
            if observed > maximum {
                return Err(invalid("src", "the image is too large"));
            }
            output.extend_from_slice(&chunk[..count]);
        }
        return Ok(output);
    }
    let file =
        std::fs::File::open(src).map_err(|source| Error::from(source).with_property("src"))?;
    let mut reader = file.take(u64::try_from(maximum).unwrap_or(u64::MAX).saturating_add(1));
    let mut output = Vec::with_capacity(maximum.min(8192));
    let mut chunk = [0_u8; 8192];
    loop {
        let count = reader
            .read(&mut chunk)
            .map_err(|source| Error::from(source).with_property("src"))?;
        if count == 0 {
            break;
        }
        let observed = output.len().saturating_add(count);
        if observed > maximum {
            return Err(invalid("src", "the image is too large"));
        }
        output.extend_from_slice(&chunk[..count]);
    }
    Ok(output)
}

/// A zero-copy filtering reader for base64 payloads.  Data URI whitespace is
/// ignored by the base64 grammar, but is never copied into a second complete
/// encoded buffer before bounded decoding.
struct Base64Input<'a> {
    bytes: &'a [u8],
    offset: usize,
}

impl<'a> Base64Input<'a> {
    fn new(bytes: &'a [u8]) -> Self {
        Self { bytes, offset: 0 }
    }
}

impl std::io::Read for Base64Input<'_> {
    fn read(&mut self, output: &mut [u8]) -> std::io::Result<usize> {
        let mut written = 0;
        while written < output.len() && self.offset < self.bytes.len() {
            let byte = self.bytes[self.offset];
            self.offset += 1;
            if byte.is_ascii_whitespace() {
                continue;
            }
            output[written] = byte;
            written += 1;
        }
        Ok(written)
    }
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
            let rect = match xlsx::Rect::from_a1(&anchor) {
                Ok(rect) => rect,
                Err(_) => {
                    // A single cell spans four columns and two rows.
                    let cell = litchi_core::selector::CellRef::parse(&anchor)
                        .ok_or_else(|| invalid("anchor", "the anchor is not an A1 range"))?;
                    return Ok(Placement::TwoCell {
                        from: Marker::cell(cell.column(), cell.row()),
                        to: Marker::cell(cell.column() + 4, cell.row() + 2),
                    });
                },
            };
            let start = rect.start();
            let (end_row, end_column) = rect.end();
            return Ok(Placement::TwoCell {
                from: Marker::cell(start.column().get(), start.row().get()),
                to: Marker::cell(end_column, end_row),
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

fn drawing_rgb(key: &'static str, color: &Color) -> Result<xlsx::Rgb> {
    match color.base() {
        crate::value::Base::Rgb(value) if color.transforms().is_empty() => {
            Ok(xlsx::Rgb::new(value.red(), value.green(), value.blue()))
        },
        _ => Err(invalid(key, "the color must be an RGB color")),
    }
}

fn parse_color(key: &'static str, value: &str, context: ColorContext) -> Result<xlsx::Rgb> {
    let color = Color::parse(value.trim(), context)
        .map_err(|_| invalid(key, "the value is not a color"))?;
    drawing_rgb(key, &color)
}

fn arg_color(key: &'static str, arg: &Arg, context: ColorContext) -> Result<xlsx::Rgb> {
    match arg {
        Arg::Color(color) => drawing_rgb(key, color),
        Arg::Text(value) => parse_color(key, value, context),
        _ => Err(invalid(key, "the value is not a color")),
    }
}

impl Xlsx {
    fn drawing_count(&self, sheet: usize, kind: Kind) -> Result<usize> {
        Ok(self.features_of(sheet, kind)?.len())
    }

    fn add_picture(
        &self,
        parent: &Parent,
        props: &Props<'_>,
        limits: &SemanticLimits,
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
        let image = read_source(&src, limits.max_staged_bytes())?;
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
        let source = self.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        let mut edit = source
            .edit_drawings_with_options(part_edit_options(&source, limits.max_snapshot_bytes()))?;
        let index = edit.add_picture(parent.sheet, &spec).map_err(authored)?;
        let commit = edit.commit()?;
        let changed = commit.diagnostics().changed();
        let native_patch_bytes = commit.patch().retained_bytes();
        let commit = self.publish_with_limit(
            source,
            commit.into_package(),
            changed,
            native_patch_bytes,
            limits.max_snapshot_bytes(),
        )?;
        Ok((commit, feature_path(parent.sheet, "picture", Some(index))))
    }

    fn add_shape(
        &self,
        parent: &Parent,
        props: &Props<'_>,
        limits: &SemanticLimits,
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
                    } else if let Some(name) = xlsx::drawing::edit::canonical(
                        xlsx::drawing::edit::SCHEME_COLORS,
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
                                xlsx::drawing::edit::canonical(
                                    xlsx::drawing::edit::DASHES,
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
        let source = self.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        let mut edit = source
            .edit_drawings_with_options(part_edit_options(&source, limits.max_snapshot_bytes()))?;
        let index = edit.add_shape(parent.sheet, &spec).map_err(authored)?;
        let commit = edit.commit()?;
        let changed = commit.diagnostics().changed();
        let native_patch_bytes = commit.patch().retained_bytes();
        let commit = self.publish_with_limit(
            source,
            commit.into_package(),
            changed,
            native_patch_bytes,
            limits.max_snapshot_bytes(),
        )?;
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
        // A new image for the picture (record 0338).
        let mut image = None;
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
                (Some(PictureKey::Src), _) => "src",
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
                "src" => {
                    let src = arg.text().ok_or_else(|| {
                        error(Code::InvalidValue, "src names an image file or a data URI")
                            .with_property("src")
                    })?;
                    let bytes = read_source(&src, limits.max_staged_bytes())?;
                    let (content_type, extension) = classify(&bytes).ok_or_else(|| {
                        error(
                            Code::UnsupportedCapability,
                            "the picture is not a PNG, JPEG, GIF, BMP or TIFF image",
                        )
                        .with_property("src")
                    })?;
                    image = Some((bytes, content_type, extension));
                },
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
        let source = self.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        let mut edit = source
            .edit_drawings_with_options(part_edit_options(&source, limits.max_snapshot_bytes()))?;
        // A `src`-only set leaves the anchor alone: no anchor rewrite.
        if image.is_none() || change != Change::default() {
            edit.set(sheet, kind, feature.index as usize, &change)
                .map_err(authored)?;
        }
        if let Some((bytes, content_type, extension)) = image {
            // Path resolution read the picture; the leaf does not read the
            // drawing again to find its image (record 0347).
            edit.replace_element_image(
                sheet,
                feature.index as usize,
                element,
                bytes,
                content_type,
                extension,
            )
            .map_err(|failure| authored(failure).with_property("src"))?;
        }
        let commit = edit.commit()?;
        let changed = commit.diagnostics().changed();
        let native_patch_bytes = commit.patch().retained_bytes();
        self.publish_with_limit(
            source,
            commit.into_package(),
            changed,
            native_patch_bytes,
            limits.max_snapshot_bytes(),
        )
    }
}

// ---------------------------------------------------------------------------
// Conditional formats

/// A CF element word or `type` value: the rule kind and the flags it implies.
struct CfType {
    kind: RuleKind,
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
        "cellis" | "highlight" => plain(RuleKind::CellIs),
        "formula" | "expression" | "formulacf" => plain(RuleKind::Expression),
        "colorscale" => plain(RuleKind::ColorScale),
        "databar" => plain(RuleKind::DataBar),
        "iconset" => plain(RuleKind::IconSet),
        "top10" | "top" | "topn" => plain(RuleKind::Top10),
        "toppercent" => CfType {
            percent: true,
            ..plain(RuleKind::Top10)
        },
        "bottom" | "bottomn" => CfType {
            bottom: true,
            ..plain(RuleKind::Top10)
        },
        "bottompercent" => CfType {
            percent: true,
            bottom: true,
            ..plain(RuleKind::Top10)
        },
        "aboveaverage" => plain(RuleKind::AboveAverage),
        "belowaverage" => CfType {
            below: true,
            ..plain(RuleKind::AboveAverage)
        },
        "duplicatevalues" => plain(RuleKind::DuplicateValues),
        "uniquevalues" => plain(RuleKind::UniqueValues),
        "containstext" | "contains" => plain(RuleKind::ContainsText),
        "notcontains" | "notcontainstext" => plain(RuleKind::NotContainsText),
        "beginswith" => plain(RuleKind::BeginsWith),
        "endswith" => plain(RuleKind::EndsWith),
        "containsblanks" => plain(RuleKind::ContainsBlanks),
        "notcontainsblanks" => plain(RuleKind::NotContainsBlanks),
        "containserrors" => plain(RuleKind::ContainsErrors),
        "notcontainserrors" => plain(RuleKind::NotContainsErrors),
        "dateoccurring" | "timeperiod" => plain(RuleKind::TimePeriod),
        _ => return None,
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

fn cf_operator(key: &'static str, value: &str) -> Result<Operator> {
    Ok(match value.trim().to_ascii_lowercase().as_str() {
        "greaterthan" | "gt" | ">" => Operator::GreaterThan,
        "lessthan" | "lt" | "<" => Operator::LessThan,
        "greaterthanorequal" | "gte" | "ge" | ">=" => Operator::GreaterThanOrEqual,
        "lessthanorequal" | "lte" | "le" | "<=" => Operator::LessThanOrEqual,
        "equal" | "eq" | "=" | "==" => Operator::Equal,
        "notequal" | "ne" | "!=" | "<>" => Operator::NotEqual,
        "between" => Operator::Between,
        "notbetween" => Operator::NotBetween,
        _ => return Err(invalid(key, "the operator is not a cellIs comparison")),
    })
}

fn cf_period(key: &'static str, value: &str) -> Result<Period> {
    let value = value.trim();
    [
        Period::Today,
        Period::Yesterday,
        Period::Tomorrow,
        Period::Last7Days,
        Period::ThisWeek,
        Period::LastWeek,
        Period::NextWeek,
        Period::ThisMonth,
        Period::LastMonth,
        Period::NextMonth,
    ]
    .into_iter()
    .find(|period| period.as_str().eq_ignore_ascii_case(value))
    .ok_or_else(|| invalid(key, "the period is not a time period"))
}

fn cf_ranges(key: &'static str, value: &str) -> Result<Vec<CfRange>> {
    let ranges: Vec<CfRange> = value
        .split_whitespace()
        .map(|range| CfRange::new(range.replace('$', "")))
        .collect::<std::result::Result<_, _>>()
        .map_err(|error| authored(error).with_property(key))?;
    if ranges.is_empty() {
        return Err(invalid(key, "the range is empty"));
    }
    Ok(ranges)
}

/// The top-left cell of the first range, the anchor of generated formulas.
fn first_cell(ranges: &[CfRange]) -> String {
    let range = ranges.first().map_or("A1", CfRange::as_str);
    range.split(':').next().unwrap_or("A1").to_owned()
}

fn quoted(text: &str) -> String {
    format!("\"{}\"", text.replace('"', "\"\""))
}

/// The formula Excel writes for a text or blank/error rule.
fn generated_formula(kind: RuleKind, text: &str, cell: &str) -> Option<String> {
    let text = quoted(text);
    Some(match kind {
        RuleKind::ContainsText => format!("NOT(ISERROR(SEARCH({text},{cell})))"),
        RuleKind::NotContainsText => format!("ISERROR(SEARCH({text},{cell}))"),
        RuleKind::BeginsWith => format!("LEFT({cell},LEN({text}))={text}"),
        RuleKind::EndsWith => format!("RIGHT({cell},LEN({text}))={text}"),
        RuleKind::ContainsBlanks => format!("LEN(TRIM({cell}))=0"),
        RuleKind::NotContainsBlanks => format!("LEN(TRIM({cell}))>0"),
        RuleKind::ContainsErrors => format!("ISERROR({cell})"),
        RuleKind::NotContainsErrors => format!("NOT(ISERROR({cell}))"),
        _ => return None,
    })
}

fn cf_threshold(key: &'static str, value: Option<&str>, fallback: ValueKind) -> Result<Threshold> {
    let value = value
        .map(str::trim)
        .filter(|value| !value.is_empty() && !value.eq_ignore_ascii_case("auto"));
    let Some(value) = value else {
        return Ok(Threshold {
            kind: fallback,
            value: None,
            formula: None,
            greater_than_or_equal: true,
        });
    };
    let (kind, text) = if let Some(rest) = value.strip_suffix('%') {
        (ValueKind::Percent, rest)
    } else if let Some((prefix, rest)) = value.split_once(':') {
        let kind = match prefix.to_ascii_lowercase().as_str() {
            "percent" => ValueKind::Percent,
            "percentile" => ValueKind::Percentile,
            "num" | "number" => ValueKind::Number,
            "formula" => ValueKind::Formula,
            _ => {
                return Err(invalid(
                    key,
                    "the threshold kind is num, percent, percentile or formula",
                ));
            },
        };
        (kind, rest)
    } else {
        (ValueKind::Number, value)
    };
    Ok(Threshold {
        kind,
        value: Some(text.to_owned()),
        formula: None,
        greater_than_or_equal: true,
    })
}

fn cf_color(rgb: xlsx::Rgb) -> CfColor {
    CfColor {
        rgb: Some(xlsx::Rgb::argb(0xFF, rgb.red(), rgb.green(), rgb.blue())),
        indexed: None,
        theme: None,
        tint: None,
        automatic: None,
    }
}

/// The parts of a CF request, applied to a fresh or an existing rule.
#[derive(Default)]
struct CfRequest {
    kind: Option<CfType>,
    ranges: Option<Vec<CfRange>>,
    stop_if_true: Option<bool>,
    operator: Option<Operator>,
    value: Option<String>,
    value2: Option<String>,
    formula: Option<String>,
    text: Option<String>,
    rank: Option<u32>,
    percent: Option<bool>,
    bottom: Option<bool>,
    above_average: Option<bool>,
    equal_average: Option<bool>,
    std_dev: Option<i32>,
    period: Option<Period>,
    color: Option<xlsx::Rgb>,
    min: Option<String>,
    max: Option<String>,
    min_length: Option<u32>,
    max_length: Option<u32>,
    show_value: Option<bool>,
    min_color: Option<xlsx::Rgb>,
    mid_color: Option<xlsx::Rgb>,
    max_color: Option<xlsx::Rgb>,
    midpoint: Option<String>,
    iconset: Option<IconSet>,
    reverse: Option<bool>,
    style: DifferentialStyle,
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
                        .filter(|rank| *rank >= 1)
                        .ok_or_else(|| invalid(name, "the rank is at least 1"))?,
                );
            },
            CfKey::Percent => request.percent = Some(arg.flag()),
            CfKey::Bottom => request.bottom = Some(arg.flag()),
            CfKey::AboveAverage => request.above_average = Some(arg.flag()),
            CfKey::EqualAverage => request.equal_average = Some(arg.flag()),
            CfKey::StdDev => {
                request.std_dev = Some(
                    i32::try_from(arg.int(name)?)
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
                let value = u32::try_from(arg.int(name)?)
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
                request.iconset = Some(
                    value
                        .trim()
                        .parse::<IconSet>()
                        .map_err(|_| invalid(name, "the icon set is not a core icon set"))?,
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

/// Apply a request to `rule`: a fresh rule of the request's kind (an add)
/// or the existing first rule of a format (a set).
fn apply_cf(rule: &mut Rule, request: &CfRequest, ranges: &[CfRange], fresh: bool) -> Result<()> {
    let kind = rule
        .rule_type
        .ok_or_else(|| error(Code::UnsupportedCapability, "the rule has no type"))?;
    if let Some(value) = request.stop_if_true {
        rule.stop_if_true = value;
    }
    match kind {
        RuleKind::CellIs => {
            if let Some(operator) = request.operator {
                rule.operator = Some(operator);
            } else if fresh {
                rule.operator = Some(Operator::GreaterThan);
            }
            let mut formulas: Vec<String> = rule.formulas.iter().cloned().collect();
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
            let between = matches!(
                rule.operator,
                Some(Operator::Between | Operator::NotBetween)
            );
            if !between {
                formulas.truncate(1);
            }
            if formulas.is_empty() {
                return Err(missing("value", "a cellIs rule needs a value"));
            }
            if between && formulas.len() < 2 {
                return Err(missing("value2", "a between rule needs value2"));
            }
            rule.formulas = formulas.into_iter().collect();
        },
        RuleKind::Expression => {
            if let Some(formula) = &request.formula.clone().or_else(|| request.value.clone()) {
                rule.formulas = std::iter::once(formula.clone()).collect();
            }
            if rule.formulas.is_empty() {
                return Err(missing("formula", "a formula rule needs a formula"));
            }
        },
        RuleKind::Top10 => {
            if let Some(rank) = request
                .rank
                .or_else(|| request.value.as_deref().and_then(|v| v.parse().ok()))
            {
                rule.rank = Some(rank);
            } else if fresh {
                rule.rank = Some(10);
            }
            if let Some(value) = request.percent {
                rule.percent = value;
            }
            if let Some(value) = request.bottom {
                rule.bottom = value;
            }
        },
        RuleKind::AboveAverage => {
            if let Some(value) = request.above_average {
                rule.above_average = value;
            }
            if let Some(value) = request.equal_average {
                rule.equal_average = value;
            }
            if let Some(value) = request.std_dev {
                rule.standard_deviations = Some(value);
            }
        },
        RuleKind::ContainsText
        | RuleKind::NotContainsText
        | RuleKind::BeginsWith
        | RuleKind::EndsWith => {
            if let Some(text) = &request.text {
                rule.text = Some(text.clone());
            }
            let text = rule
                .text
                .clone()
                .ok_or_else(|| missing("text", "a text rule needs text"))?;
            rule.operator = Some(match kind {
                RuleKind::ContainsText => Operator::ContainsText,
                RuleKind::NotContainsText => Operator::NotContains,
                RuleKind::BeginsWith => Operator::BeginsWith,
                _ => Operator::EndsWith,
            });
            if fresh || request.text.is_some() || request.ranges.is_some() {
                let formula = generated_formula(kind, &text, &first_cell(ranges))
                    .ok_or_else(|| error(Code::InternalError, "no generated formula"))?;
                rule.formulas = std::iter::once(formula).collect();
            }
        },
        RuleKind::ContainsBlanks
        | RuleKind::NotContainsBlanks
        | RuleKind::ContainsErrors
        | RuleKind::NotContainsErrors => {
            if fresh || request.ranges.is_some() {
                let formula = generated_formula(kind, "", &first_cell(ranges))
                    .ok_or_else(|| error(Code::InternalError, "no generated formula"))?;
                rule.formulas = std::iter::once(formula).collect();
            }
        },
        RuleKind::TimePeriod => {
            if let Some(period) = request.period {
                rule.time_period = Some(period);
            } else if fresh {
                rule.time_period = Some(Period::Today);
            }
            if fresh || request.period.is_some() || request.ranges.is_some() {
                let cell = first_cell(ranges);
                let formula = match rule.time_period.unwrap_or(Period::Today) {
                    Period::Today => format!("FLOOR({cell},1)=TODAY()"),
                    Period::Yesterday => format!("FLOOR({cell},1)=TODAY()-1"),
                    Period::Tomorrow => format!("FLOOR({cell},1)=TODAY()+1"),
                    Period::Last7Days => {
                        format!("AND(TODAY()-FLOOR({cell},1)<=6,FLOOR({cell},1)<=TODAY())")
                    },
                    Period::ThisWeek => format!(
                        "AND(TODAY()-ROUNDDOWN({cell},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN({cell},0)-TODAY()<=7-WEEKDAY(TODAY()))"
                    ),
                    Period::LastWeek => format!(
                        "AND(TODAY()-ROUNDDOWN({cell},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN({cell},0)<(WEEKDAY(TODAY())+7))"
                    ),
                    Period::NextWeek => format!(
                        "AND(ROUNDDOWN({cell},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN({cell},0)-TODAY()<(15-WEEKDAY(TODAY())))"
                    ),
                    Period::ThisMonth => {
                        format!("AND(MONTH({cell})=MONTH(TODAY()),YEAR({cell})=YEAR(TODAY()))")
                    },
                    Period::LastMonth => format!(
                        "AND(MONTH({cell})=MONTH(EDATE(TODAY(),0-1)),YEAR({cell})=YEAR(EDATE(TODAY(),0-1)))"
                    ),
                    _ => format!(
                        "AND(MONTH({cell})=MONTH(EDATE(TODAY(),0+1)),YEAR({cell})=YEAR(EDATE(TODAY(),0+1)))"
                    ),
                };
                rule.formulas = std::iter::once(formula).collect();
            }
        },
        RuleKind::DataBar => {
            let mut bar = match rule.payload.take() {
                Some(Payload::DataBar(bar)) => bar,
                _ => DataBar {
                    thresholds: vec![
                        cf_threshold("min", None, ValueKind::Min)?,
                        cf_threshold("max", None, ValueKind::Max)?,
                    ],
                    colors: vec![NamedColor {
                        role: ColorRole::Color,
                        color: cf_color(xlsx::Rgb::new(0x63, 0x8E, 0xC6)),
                    }],
                    min_length: 10,
                    max_length: 90,
                    show_value: true,
                    border: false,
                    gradient: true,
                    direction: Direction::Context,
                    axis_position: xlsx::conditional_formatting::Axis::Automatic,
                },
            };
            if request.min.is_some() {
                bar.thresholds[0] = cf_threshold("min", request.min.as_deref(), ValueKind::Min)?;
            }
            if request.max.is_some() {
                bar.thresholds[1] = cf_threshold("max", request.max.as_deref(), ValueKind::Max)?;
            }
            if let Some(color) = request.color {
                bar.colors = vec![NamedColor {
                    role: ColorRole::Color,
                    color: cf_color(color),
                }];
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
            rule.payload = Some(Payload::DataBar(bar));
        },
        RuleKind::ColorScale => {
            let mut scale = match rule.payload.take() {
                Some(Payload::ColorScale(scale)) => scale,
                _ => ColorScale {
                    thresholds: vec![
                        cf_threshold("min", None, ValueKind::Min)?,
                        cf_threshold("max", None, ValueKind::Max)?,
                    ],
                    colors: vec![
                        cf_color(xlsx::Rgb::new(0xF8, 0x69, 0x6B)),
                        cf_color(xlsx::Rgb::new(0x63, 0xBE, 0x7B)),
                    ],
                },
            };
            if request.mid_color.is_some() || request.midpoint.is_some() {
                if scale.colors.len() == 2 {
                    let mid = Threshold {
                        kind: ValueKind::Percentile,
                        value: Some("50".to_owned()),
                        formula: None,
                        greater_than_or_equal: true,
                    };
                    scale.thresholds.insert(1, mid);
                    scale
                        .colors
                        .insert(1, cf_color(xlsx::Rgb::new(0xFF, 0xEB, 0x84)));
                }
                if let Some(midpoint) = &request.midpoint {
                    let mut threshold =
                        cf_threshold("midpoint", Some(midpoint), ValueKind::Percentile)?;
                    if threshold.kind == ValueKind::Number && !midpoint.contains(':') {
                        threshold.kind = ValueKind::Percentile;
                    }
                    scale.thresholds[1] = threshold;
                }
                if let Some(color) = request.mid_color {
                    scale.colors[1] = cf_color(color);
                }
            }
            if let Some(color) = request.min_color {
                scale.colors[0] = cf_color(color);
            }
            if let Some(color) = request.max_color {
                let last = scale.colors.len() - 1;
                scale.colors[last] = cf_color(color);
            }
            rule.payload = Some(Payload::ColorScale(scale));
        },
        RuleKind::IconSet => {
            let mut icons = match rule.payload.take() {
                Some(Payload::IconSet(icons)) => icons,
                _ => Icons {
                    set: IconSet::ThreeTrafficLights1,
                    thresholds: Vec::new(),
                    show_value: true,
                    percent: true,
                    reverse: false,
                },
            };
            if let Some(set) = request.iconset {
                icons.set = set;
                icons.thresholds.clear();
            }
            if icons.thresholds.is_empty() {
                let count = u32::from(icons.set.len());
                icons.thresholds = (0..count)
                    .map(|step| Threshold {
                        kind: ValueKind::Percent,
                        value: Some((step * 100 / count).to_string()),
                        formula: None,
                        greater_than_or_equal: true,
                    })
                    .collect();
            }
            if let Some(value) = request.reverse {
                icons.reverse = value;
            }
            if let Some(value) = request.show_value {
                icons.show_value = value;
            }
            rule.payload = Some(Payload::IconSet(icons));
        },
        RuleKind::DuplicateValues | RuleKind::UniqueValues => {},
    }
    Ok(())
}

impl Xlsx {
    fn add_cf(
        &self,
        parent: &Parent,
        implied: Option<&'static str>,
        props: &Props<'_>,
        limits: &SemanticLimits,
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
        let source = self.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        let mut edit = source.edit_conditional_formats_with_options(part_edit_options(
            &source,
            limits.max_snapshot_bytes(),
        ))?;
        let mut formats = edit.formats(parent.sheet)?;
        let priority = formats
            .iter()
            .flat_map(|format| format.rules.iter().filter_map(|rule| rule.priority))
            .max()
            .unwrap_or(0)
            .checked_add(1)
            .ok_or_else(|| error(Code::OutOfRange, "the priorities are exhausted"))?;
        let mut rule = Rule::new(kind.kind, priority).map_err(authored)?;
        if kind.percent {
            rule.percent = true;
        }
        if kind.bottom {
            rule.bottom = true;
        }
        if kind.below {
            rule.above_average = false;
        }
        apply_cf(&mut rule, &request, &ranges, true)?;
        if !request.style.is_empty() {
            let dxf = edit.add_differential(&request.style).map_err(authored)?;
            rule.differential_format = Some(DifferentialRef::StylesIndex(dxf));
        }
        formats.push(Formatting::new(ranges, vec![rule]).map_err(authored)?);
        let index = formats.len() - 1;
        edit.set_formats(parent.sheet, formats).map_err(authored)?;
        let commit = edit.commit()?;
        let changed = commit.diagnostics().changed();
        let native_patch_bytes = commit.patch().retained_bytes();
        let commit = self.publish_with_limit(
            source,
            commit.into_package(),
            changed,
            native_patch_bytes,
            limits.max_snapshot_bytes(),
        )?;
        Ok((commit, feature_path(parent.sheet, "cf", Some(index))))
    }

    fn set_cf(
        &self,
        sheet: usize,
        feature: &Feature,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        let request = cf_request(Operation::Set, props)?;
        let source = self.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        let mut edit = source.edit_conditional_formats_with_options(part_edit_options(
            &source,
            limits.max_snapshot_bytes(),
        ))?;
        let mut formats = edit.formats(sheet)?;
        let format = formats
            .get_mut(feature.index as usize)
            .ok_or_else(|| error(Code::PathNotFound, "the conditional format is missing"))?;
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
        if !request.style.is_empty() {
            let current = match (&feature.data, &rule.differential_format) {
                (Data::Cf(cf), Some(DifferentialRef::StylesIndex(_))) => {
                    cf.dxf.clone().unwrap_or_default()
                },
                _ => DifferentialStyle::default(),
            };
            let merged = merge_style(current, &request.style);
            let dxf = edit.add_differential(&merged).map_err(authored)?;
            let rule = formats[feature.index as usize]
                .rules
                .first_mut()
                .ok_or_else(|| error(Code::InternalError, "the rule vanished"))?;
            rule.differential_format = Some(DifferentialRef::StylesIndex(dxf));
        }
        edit.set_formats(sheet, formats).map_err(authored)?;
        let commit = edit.commit()?;
        let changed = commit.diagnostics().changed();
        let native_patch_bytes = commit.patch().retained_bytes();
        self.publish_with_limit(
            source,
            commit.into_package(),
            changed,
            native_patch_bytes,
            limits.max_snapshot_bytes(),
        )
    }
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

// ---------------------------------------------------------------------------
// Data validations

fn locate(collections: &[xlsx::data_validation::Collection], index: u32) -> Result<(usize, usize)> {
    let mut remaining = index as usize;
    for (position, collection) in collections.iter().enumerate() {
        let count = collection.validations().len();
        if remaining < count {
            return Ok((position, remaining));
        }
        remaining -= count;
    }
    Err(error(Code::PathNotFound, "the data validation is missing"))
}

/// The rectangles of a validation's cells, for the overlap check.
fn rects(sqref: &xlsx::data_validation::Sqref) -> Vec<xlsx::Rect> {
    sqref
        .ranges()
        .iter()
        .filter_map(|range| {
            let text = range.as_str();
            xlsx::Rect::from_a1(text).ok().or_else(|| {
                let cell = litchi_core::selector::CellRef::parse(text)?;
                let at = xlsx::Address::at(cell.row(), cell.column()).ok()?;
                Some(xlsx::Rect::single(at))
            })
        })
        .collect()
}

fn overlaps(left: &[xlsx::Rect], right: &[xlsx::Rect]) -> bool {
    left.iter().any(|a| {
        right.iter().any(|b| {
            let (a_rows, a_cols) = a.end();
            let (b_rows, b_cols) = b.end();
            a.start().row().get() < b_rows
                && b.start().row().get() < a_rows
                && a.start().column().get() < b_cols
                && b.start().column().get() < a_cols
        })
    })
}

/// A validation formula: a leading `=` is dropped, and a list's literal
/// options are quoted.
fn validation_formula(
    key: &'static str,
    value: &str,
    list: bool,
) -> Result<xlsx::data_validation::ListSource> {
    use xlsx::data_validation::{Formula, ListSource};
    let value = value.trim();
    if value.len() > 255 {
        return Err(invalid(
            key,
            "a validation formula holds at most 255 characters",
        ));
    }
    let formula = if let Some(rest) = value.strip_prefix('=') {
        rest.to_owned()
    } else if list && !value.starts_with('"') {
        if value.contains('"') {
            return Err(invalid(key, "list options cannot contain quotes"));
        }
        quoted(value)
    } else {
        value.to_owned()
    };
    Ok(ListSource::Formula(
        Formula::new(formula).map_err(|error| authored(error).with_property(key))?,
    ))
}

fn apply_validation(
    validation: &mut xlsx::data_validation::Validation,
    props: &Props<'_>,
    operation: Operation,
) -> Result<()> {
    use xlsx::data_validation::{
        ListSource, Sqref, ValidationErrorStyle, ValidationOperator, ValidationType,
    };
    let admitted = admit(&features::VALIDATION_TABLE, operation, props)?;
    // The type first: a list's formula is quoted by type.
    for (key, name, arg) in &admitted {
        if *key == ValidationKey::Type {
            let value = arg.text().unwrap_or_default();
            let value = match value.to_ascii_lowercase().as_str() {
                "textlength" => "textLength".to_owned(),
                "any" => "none".to_owned(),
                other => other.to_owned(),
            };
            validation.set_validation_type(
                ValidationType::parse(&value)
                    .map_err(|error| authored(error).with_property(*name))?,
            );
        }
    }
    let list = validation.validation_type() == ValidationType::List;
    for (key, name, arg) in admitted {
        let text = || arg.text().unwrap_or_default();
        let wrap =
            |result: xlsx::Result<()>| result.map_err(|error| authored(error).with_property(name));
        match key {
            ValidationKey::Type => {},
            ValidationKey::Ref => validation.set_sqref(
                Sqref::parse(text().replace('$', ""))
                    .map_err(|error| authored(error).with_property(name))?,
            ),
            ValidationKey::Operator => {
                let value = text();
                let operator = [
                    ValidationOperator::Between,
                    ValidationOperator::NotBetween,
                    ValidationOperator::Equal,
                    ValidationOperator::NotEqual,
                    ValidationOperator::LessThan,
                    ValidationOperator::LessThanOrEqual,
                    ValidationOperator::GreaterThan,
                    ValidationOperator::GreaterThanOrEqual,
                ]
                .into_iter()
                .find(|operator| operator.as_str().eq_ignore_ascii_case(value.trim()))
                .ok_or_else(|| invalid(name, "the operator is not a validation comparison"))?;
                validation.set_operator(operator);
            },
            ValidationKey::Formula1 => wrap(
                validation.set_formula1(
                    arg.text()
                        .map(|value| validation_formula(name, &value, list))
                        .transpose()?,
                ),
            )?,
            ValidationKey::Formula2 => wrap(
                validation.set_formula2(
                    arg.text()
                        .map(|value| {
                            validation_formula(name, &value, false).map(|source| match source {
                                ListSource::Formula(formula) => formula,
                                ListSource::QuotedList(_) => {
                                    unreachable!("formula2 is never a list")
                                },
                            })
                        })
                        .transpose()?,
                ),
            )?,
            ValidationKey::AllowBlank => validation.set_allow_blank(arg.flag()),
            ValidationKey::ShowError => validation.set_show_error_message(arg.flag()),
            ValidationKey::ShowInput => validation.set_show_input_message(arg.flag()),
            ValidationKey::ErrorTitle => wrap(validation.set_error_title(arg.text()))?,
            ValidationKey::Error => wrap(validation.set_error(arg.text()))?,
            ValidationKey::PromptTitle => wrap(validation.set_prompt_title(arg.text()))?,
            ValidationKey::Prompt => wrap(validation.set_prompt(arg.text()))?,
            ValidationKey::ErrorStyle => {
                let value = text().to_ascii_lowercase();
                let value = match value.as_str() {
                    "warn" => "warning",
                    "info" => "information",
                    other => other,
                };
                validation.set_error_style(
                    ValidationErrorStyle::parse(value)
                        .map_err(|error| authored(error).with_property(name))?,
                );
            },
            ValidationKey::InCellDropdown => validation.set_show_drop_down(!arg.flag()),
        }
    }
    validation.validate().map_err(authored)
}

impl Xlsx {
    fn add_validation(
        &self,
        parent: &Parent,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        use xlsx::data_validation::{Collection, Source, Sqref, Validation, ValidationType};
        let reference = props
            .iter()
            .find(|(key, _)| {
                ["ref", "sqref", "range"]
                    .iter()
                    .any(|name| key.as_str().eq_ignore_ascii_case(name))
            })
            .is_some();
        let fallback = match (&parent.reference, reference) {
            (_, true) => "A1".to_owned(),
            (Some(reference), false) => reference.clone(),
            (None, false) => {
                return Err(missing("ref", "a data validation needs a ref"));
            },
        };
        let mut validation = Validation::new(
            Source::Core,
            ValidationType::None,
            Sqref::parse(&fallback).map_err(|error| authored(error).with_property("ref"))?,
        );
        validation.set_allow_blank(true);
        validation.set_show_error_message(true);
        validation.set_show_input_message(true);
        apply_validation(&mut validation, props, Operation::Add)?;
        let source = self.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        let mut edit = source.edit_data_validations_with_options(part_edit_options(
            &source,
            limits.max_snapshot_bytes(),
        ))?;
        let mut collections = edit.validations(parent.sheet)?;
        let added = rects(validation.sqref());
        if collections
            .iter()
            .flat_map(|collection| collection.validations())
            .any(|existing| overlaps(&rects(existing.sqref()), &added))
        {
            return Err(error(
                Code::InvalidArgument,
                "the range overlaps an existing data validation",
            )
            .with_property("ref"));
        }
        let index = collections
            .iter()
            .map(|collection| collection.validations().len())
            .sum::<usize>();
        match collections
            .iter_mut()
            .rposition(|collection| collection.source() == Source::Core)
        {
            Some(position) if position + 1 == collections.len() => {
                let collection = &mut collections[position];
                let mut validations = collection.validations().to_vec();
                validations.push(validation);
                collection.set_validations(validations).map_err(authored)?;
            },
            _ => {
                collections.push(Collection::new(Source::Core, vec![validation]).map_err(authored)?)
            },
        }
        edit.set_validations(parent.sheet, collections)
            .map_err(authored)?;
        let commit = edit.commit()?;
        let changed = commit.diagnostics().changed();
        let native_patch_bytes = commit.patch().retained_bytes();
        let commit = self.publish_with_limit(
            source,
            commit.into_package(),
            changed,
            native_patch_bytes,
            limits.max_snapshot_bytes(),
        )?;
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
        let source = self.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        let mut edit = source.edit_data_validations_with_options(part_edit_options(
            &source,
            limits.max_snapshot_bytes(),
        ))?;
        let mut collections = edit.validations(sheet)?;
        let (collection, item) = locate(&collections, feature.index)?;
        let mut validations = collections[collection].validations().to_vec();
        apply_validation(&mut validations[item], props, Operation::Set)?;
        let changed_rects = rects(validations[item].sqref());
        let clash = collections.iter().enumerate().any(|(position, held)| {
            held.validations().iter().enumerate().any(|(index, other)| {
                (position, index) != (collection, item)
                    && overlaps(&rects(other.sqref()), &changed_rects)
            })
        });
        if clash {
            return Err(error(
                Code::InvalidArgument,
                "the range overlaps another data validation",
            )
            .with_property("ref"));
        }
        collections[collection]
            .set_validations(validations)
            .map_err(authored)?;
        edit.set_validations(sheet, collections).map_err(authored)?;
        let commit = edit.commit()?;
        let changed = commit.diagnostics().changed();
        let native_patch_bytes = commit.patch().retained_bytes();
        self.publish_with_limit(
            source,
            commit.into_package(),
            changed,
            native_patch_bytes,
            limits.max_snapshot_bytes(),
        )
    }
}

// ---------------------------------------------------------------------------
// The auto-filter

fn filter_range(key: &'static str, value: &str) -> Result<xlsx::auto_filter::Definition> {
    let range = xlsx::auto_filter::Range::new(value.trim().replace('$', ""))
        .map_err(|error| authored(error).with_property(key))?;
    Ok(xlsx::auto_filter::Definition::new(Some(range)))
}

/// A `criteriaN.OP` key: the filter column (an offset from the range's
/// first column) and the lowercased operator.
fn criteria_key(key: &str) -> Option<(u32, String)> {
    let rest = key.strip_prefix("criteria")?;
    let (column, operator) = rest.split_once('.')?;
    if column.is_empty()
        || !column.bytes().all(|byte| byte.is_ascii_digit())
        || operator.is_empty()
        || !operator.bytes().all(|byte| byte.is_ascii_alphabetic())
    {
        return None;
    }
    Some((column.parse().ok()?, operator.to_ascii_lowercase()))
}

fn truthy(value: &str) -> bool {
    matches!(
        value.trim().to_ascii_lowercase().as_str(),
        "true" | "1" | "yes" | "on"
    )
}

/// The OfficeCLI dynamic-filter token of a type.
fn dynamic_type(value: &str) -> Option<xlsx::auto_filter::DynamicType> {
    use xlsx::auto_filter::DynamicType as T;
    Some(match value.trim() {
        "aboveAverage" => T::AboveAverage,
        "belowAverage" => T::BelowAverage,
        "tomorrow" => T::Tomorrow,
        "today" => T::Today,
        "yesterday" => T::Yesterday,
        "nextWeek" => T::NextWeek,
        "thisWeek" => T::ThisWeek,
        "lastWeek" => T::LastWeek,
        "nextMonth" => T::NextMonth,
        "thisMonth" => T::ThisMonth,
        "lastMonth" => T::LastMonth,
        "nextQuarter" => T::NextQuarter,
        "thisQuarter" => T::ThisQuarter,
        "lastQuarter" => T::LastQuarter,
        "nextYear" => T::NextYear,
        "thisYear" => T::ThisYear,
        "lastYear" => T::LastYear,
        "yearToDate" => T::YearToDate,
        "Q1" => T::Q1,
        "Q2" => T::Q2,
        "Q3" => T::Q3,
        "Q4" => T::Q4,
        "M1" => T::M1,
        "M2" => T::M2,
        "M3" => T::M3,
        "M4" => T::M4,
        "M5" => T::M5,
        "M6" => T::M6,
        "M7" => T::M7,
        "M8" => T::M8,
        "M9" => T::M9,
        "M10" => T::M10,
        "M11" => T::M11,
        "M12" => T::M12,
        _ => return None,
    })
}

/// The filter columns of an auto-filter add's `criteriaN.OP=VAL`
/// properties (OfficeCLI's AF1): one column per distinct N, its custom
/// filters (`equals`, `notEquals`, `contains`, `doesNotContain`,
/// `beginsWith`, `endsWith`, `gt`, `gte`, `lt`, `lte`, `between`,
/// `notBetween`, `nonBlanks`) unless a dedicated operator (`top`,
/// `topPercent`, `bottom`, `bottomPercent`, `blanks`, `values`, `dynamic`)
/// names the column's filter.
fn filter_columns(criteria: Vec<(u32, String, String)>) -> Result<Vec<xlsx::auto_filter::Column>> {
    use xlsx::auto_filter::{
        Calendar, Column, Custom, Customs, Dynamic, Item, Operator, Payload, Top10, Values,
    };
    let mut grouped: std::collections::BTreeMap<u32, Vec<(String, String)>> =
        std::collections::BTreeMap::new();
    for (column, operator, value) in criteria {
        grouped.entry(column).or_default().push((operator, value));
    }
    let mut columns = Vec::with_capacity(grouped.len());
    for (column_id, entries) in grouped {
        let key = "criteria0";
        let mut column =
            Column::new(column_id).map_err(|error| authored(error).with_property(key))?;
        let mut customs: Vec<(Operator, String)> = Vec::new();
        let mut and = false;
        let mut dedicated: Option<Payload> = None;
        let mut values: Option<(bool, Vec<Item>)> = None;
        for (operator, raw) in entries {
            match operator.as_str() {
                "equals" => customs.push((Operator::Equal, raw)),
                "notequals" => customs.push((Operator::NotEqual, raw)),
                "contains" | "doesnotcontain" => {
                    let wild = if raw.contains('*') {
                        raw
                    } else {
                        format!("*{raw}*")
                    };
                    customs.push((
                        if operator == "contains" {
                            Operator::Equal
                        } else {
                            Operator::NotEqual
                        },
                        wild,
                    ));
                },
                "beginswith" => {
                    let wild = if raw.ends_with('*') {
                        raw
                    } else {
                        format!("{raw}*")
                    };
                    customs.push((Operator::Equal, wild));
                },
                "endswith" => {
                    let wild = if raw.starts_with('*') {
                        raw
                    } else {
                        format!("*{raw}")
                    };
                    customs.push((Operator::Equal, wild));
                },
                "gt" => customs.push((Operator::GreaterThan, raw)),
                "gte" => customs.push((Operator::GreaterThanOrEqual, raw)),
                "lt" => customs.push((Operator::LessThan, raw)),
                "lte" => customs.push((Operator::LessThanOrEqual, raw)),
                "between" | "notbetween" => {
                    let Some((low, high)) = raw.split_once(',') else {
                        return Err(invalid(key, "between and notBetween need 'lo,hi'"));
                    };
                    let (low, high) = (low.trim().to_owned(), high.trim().to_owned());
                    if operator == "between" {
                        customs.push((Operator::GreaterThanOrEqual, low));
                        customs.push((Operator::LessThanOrEqual, high));
                        and = true;
                    } else {
                        customs.push((Operator::LessThan, low));
                        customs.push((Operator::GreaterThan, high));
                    }
                },
                "top" | "toppercent" | "bottom" | "bottompercent" => {
                    let value: f64 = raw
                        .trim()
                        .parse()
                        .map_err(|_| invalid(key, "a top or bottom criterion needs a number"))?;
                    let top = operator.starts_with("top");
                    let percent = operator.ends_with("percent");
                    dedicated = Some(Payload::Top10(
                        Top10::new(top, percent, value, None)
                            .map_err(|error| authored(error).with_property(key))?,
                    ));
                },
                "blanks" => {
                    if truthy(&raw) {
                        values.get_or_insert_with(|| (false, Vec::new())).0 = true;
                    }
                },
                "nonblanks" => {
                    if truthy(&raw) {
                        customs.push((Operator::NotEqual, String::new()));
                    }
                },
                "values" => {
                    let items = &mut values.get_or_insert_with(|| (false, Vec::new())).1;
                    items.extend(
                        raw.split(',')
                            .map(str::trim)
                            .filter(|value| !value.is_empty())
                            .map(|value| Item::Value(value.to_owned())),
                    );
                },
                "dynamic" => {
                    let filter_type = dynamic_type(&raw).ok_or_else(|| {
                        invalid(key, "the dynamic filter type is not one Excel knows")
                    })?;
                    dedicated = Some(Payload::Dynamic(
                        Dynamic::new(filter_type, None, None)
                            .map_err(|error| authored(error).with_property(key))?,
                    ));
                },
                _ => {
                    return Err(invalid(
                        key,
                        "the criterion operator is equals, notEquals, contains, doesNotContain, beginsWith, endsWith, gt, gte, lt, lte, between, notBetween, top, topPercent, bottom, bottomPercent, blanks, nonBlanks, values or dynamic",
                    ));
                },
            }
        }
        let payload = match (dedicated, values) {
            (Some(payload), _) => Some(payload),
            (None, Some((blank, items))) => Some(Payload::Values(
                Values::new(blank, Calendar::None, items)
                    .map_err(|error| authored(error).with_property(key))?,
            )),
            (None, None) if !customs.is_empty() => {
                let filters = customs
                    .into_iter()
                    .map(|(operator, value)| Custom::new(operator, value))
                    .collect::<xlsx::Result<Vec<_>>>()
                    .map_err(|error| authored(error).with_property(key))?;
                Some(Payload::Custom(
                    Customs::new(and, filters)
                        .map_err(|error| authored(error).with_property(key))?,
                ))
            },
            (None, None) => None,
        };
        column.set_payload(payload);
        columns.push(column);
    }
    Ok(columns)
}

impl Xlsx {
    fn add_autofilter(
        &self,
        parent: &Parent,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        let mut range = parent.reference.clone();
        // `criteriaN.OP=VAL` keys are a family the registry names by its
        // first member (`criteria0`); they are admitted here.
        let mut criteria = Vec::new();
        let mut rest = Props::with_capacity(props.len());
        for (key, value) in props.iter() {
            match criteria_key(key.as_str()) {
                Some((column, operator)) => {
                    let text = match value {
                        crate::props::PropValue::Text(text) => text.to_string(),
                        crate::props::PropValue::Int(number) => number.to_string(),
                        crate::props::PropValue::Number(number) => {
                            crate::value::DotNetDouble(*number).to_string()
                        },
                        crate::props::PropValue::Bool(flag) => flag.to_string(),
                        _ => {
                            return Err(invalid(
                                "criteria0",
                                "a criterion value is text or a number",
                            ));
                        },
                    };
                    criteria.push((column, operator, text));
                },
                None => rest.push(key.reborrow(), value.reborrow()),
            }
        }
        for (key, _, arg) in admit(&features::AUTOFILTER_TABLE, Operation::Add, &rest)? {
            match key {
                AutofilterKey::Range => range = arg.text(),
                AutofilterKey::Criteria => {
                    return Err(invalid(
                        "criteria0",
                        "a column criterion is written criteriaN.OP=VAL",
                    )
                    .with_suggestion("for example criteria0.equals=Red"));
                },
            }
        }
        let range = range.ok_or_else(|| missing("range", "an auto-filter needs a range"))?;
        let mut definition = filter_range("range", &range)?;
        definition.columns = filter_columns(criteria)?;
        let source = self.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        let mut edit = source.edit_auto_filter_with_options(part_edit_options(
            &source,
            limits.max_snapshot_bytes(),
        ))?;
        edit.set_auto_filter(parent.sheet, Some(definition))
            .map_err(authored)?;
        let commit = edit.commit()?;
        let changed = commit.diagnostics().changed();
        let native_patch_bytes = commit.patch().retained_bytes();
        let commit = self.publish_with_limit(
            source,
            commit.into_package(),
            changed,
            native_patch_bytes,
            limits.max_snapshot_bytes(),
        )?;
        Ok((commit, feature_path(parent.sheet, "autofilter", None)))
    }

    fn set_autofilter(
        &self,
        sheet: usize,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        let source = self.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        let mut edit = source.edit_auto_filter_with_options(part_edit_options(
            &source,
            limits.max_snapshot_bytes(),
        ))?;
        let mut definition = edit
            .auto_filter(sheet)?
            .ok_or_else(|| error(Code::PathNotFound, "the worksheet has no auto-filter"))?;
        for (key, name, arg) in admit(&features::AUTOFILTER_TABLE, Operation::Set, props)? {
            if key == AutofilterKey::Range {
                let replacement = filter_range(name, &arg.text().unwrap_or_default())?;
                definition.reference = replacement.reference;
            }
        }
        edit.set_auto_filter(sheet, Some(definition))?;
        let commit = edit.commit()?;
        let changed = commit.diagnostics().changed();
        let native_patch_bytes = commit.patch().retained_bytes();
        self.publish_with_limit(
            source,
            commit.into_package(),
            changed,
            native_patch_bytes,
            limits.max_snapshot_bytes(),
        )
    }
}

// ---------------------------------------------------------------------------
// Page breaks

fn break_index(kind: Kind, key: &'static str, arg: &Arg) -> Result<u32> {
    let (maximum, letters) = if kind == Kind::RowBreak {
        (1_048_576_i64, false)
    } else {
        (16_384, true)
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

impl Xlsx {
    fn breaks(&self, sheet: usize, kind: Kind) -> Result<Vec<xlsx::page_breaks::Break>> {
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
        source: xlsx::Package,
        sheet: usize,
        kind: Kind,
        breaks: Vec<xlsx::page_breaks::Break>,
        max_snapshot_bytes: usize,
    ) -> Result<Commit> {
        let mut edit = source
            .edit_page_breaks_with_options(sheet, part_edit_options(&source, max_snapshot_bytes))?;
        let changed = if breaks.is_empty() {
            if kind == Kind::RowBreak {
                edit.remove_horizontal()?
            } else {
                edit.remove_vertical()?
            }
        } else if kind == Kind::RowBreak {
            edit.set_horizontal(
                xlsx::page_breaks::Collection::horizontal(breaks).map_err(authored)?,
            )?
        } else {
            edit.set_vertical(xlsx::page_breaks::Collection::vertical(breaks).map_err(authored)?)?
        };
        let committed = edit.commit()?;
        let changed = changed && committed.diagnostics().changed();
        let native_patch_bytes = committed.patch().retained_bytes();
        self.publish_with_limit(
            source,
            committed.into_package(),
            changed,
            native_patch_bytes,
            max_snapshot_bytes,
        )
    }

    fn break_spec(
        kind: Kind,
        current: Option<xlsx::page_breaks::Break>,
        props: &Props<'_>,
        operation: Operation,
    ) -> Result<xlsx::page_breaks::Break> {
        let full = if kind == Kind::RowBreak {
            16_383
        } else {
            1_048_575
        };
        let mut id = current.map(xlsx::page_breaks::Break::id);
        let mut minimum = current.map_or(0, xlsx::page_breaks::Break::minimum);
        let mut maximum = current.map_or(full, xlsx::page_breaks::Break::maximum);
        let mut manual = current.is_none_or(xlsx::page_breaks::Break::is_manual);
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
        Ok(xlsx::page_breaks::Break::new(id, minimum, maximum)
            .map_err(|error| authored(error).with_property(key))?
            .with_manual(manual))
    }

    fn add_break(
        &self,
        parent: &Parent,
        kind: Kind,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        let item = Self::break_spec(kind, None, props, Operation::Add)?;
        let mut breaks = self.breaks(parent.sheet, kind)?;
        if breaks.iter().any(|held| held.id() == item.id()) {
            return Err(error(
                Code::InvalidArgument,
                "the worksheet already has this page break",
            ));
        }
        let index = breaks.partition_point(|held| held.id() < item.id());
        breaks.insert(index, item);
        let source = self.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        let commit = self.publish_breaks(
            source,
            parent.sheet,
            kind,
            breaks,
            limits.max_snapshot_bytes(),
        )?;
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
        if breaks.iter().any(|held| held.id() == item.id()) {
            return Err(error(
                Code::InvalidArgument,
                "the worksheet already has this page break",
            ));
        }
        let index = breaks.partition_point(|held| held.id() < item.id());
        breaks.insert(index, item);
        let source = self.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        self.publish_breaks(
            source,
            sheet,
            feature.kind,
            breaks,
            limits.max_snapshot_bytes(),
        )
    }
}
