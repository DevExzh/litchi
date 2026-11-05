//! Add, set and remove of ODS worksheet features (record 0327).
//!
//! Every verb admits its properties through the kind's registry table in
//! `features` (the XLSX vocabulary of record 0320), builds the leaf value
//! and stages one leaf `Edit`: hyperlinks are `text:a` anchors of a cell
//! (the worksheet family), the other kinds provenance splices of
//! `content.xml` (and `styles.xml` for a conditional format's cell style)
//! through the `sheet_features` verbs, so every element they do not touch
//! keeps its bytes.
//!
//! Formulas (CF values and rules, validation formulas) are written in A1
//! syntax without a leading `=` and converted to OpenFormula references
//! (`[.A1]`). Drawings are placed on the sheet in lengths (`2cm`,
//! `914400emu`); cell counts are refused.

use std::collections::BTreeMap;

use litchi_core::FileFormat;
use litchi_core::selector::{Path, Predicate, Step};
use litchi_ods::advanced::sheet_features::{
    self as sf, DrawingChange, DrawingKind, DrawingSpec, FormatStyle, Rect, ValidationSpec,
    address, to_odff,
};
use litchi_ods::document::{Collision, Resource};
use litchi_ods::model::conditional_format::{self as cf, EntryType};
use litchi_ods::model::hyperlink::Link as LeafLink;
use litchi_ods::{Cell, CellValue};

use super::Ods;
use super::features::{
    self, AutofilterKey, BreakKey, CF_OPERATORS, CfKey, CfKind, Data, Feature, GEOMETRIES,
    HyperlinkKey, PERIODS, PictureKey, ShapeKey, VALIDATION_OPERATORS, ValidationKey,
};
use super::node::{Item, Node};
use super::spec::Kind;
use super::write::{Commit, finish, leaf, odf_leaf, stage_cells};
use crate::error::{Code, Error, Result};
use crate::new_node::NewNode;
use crate::placement::Placement as NodePlacement;
use crate::props::Props;
use crate::registry::{Operation, Table};
use crate::target::TargetArg;
use crate::transaction::SemanticLimits;
use crate::value::{Color, ColorContext, Length, Scalar};

fn error(code: Code, message: impl Into<String>) -> Error {
    Error::new(code, message.into()).with_format(FileFormat::Ods)
}

fn invalid(key: &'static str, message: impl Into<String>) -> Error {
    error(Code::InvalidValue, message).with_property(key)
}

fn missing(key: &'static str, message: &'static str) -> Error {
    error(Code::MissingProperty, message).with_property(key)
}

/// A leaf error: a generic unsupported shape is a capability refusal, while a
/// typed edit refusal keeps the `RefusalKind` supplied by the leaf.
fn native(error: litchi_odf_common::Error) -> Error {
    let code = if matches!(
        &error,
        litchi_odf_common::Error::Core(litchi_core::Error::Unsupported(_))
    ) {
        Code::UnsupportedCapability
    } else {
        Code::InvalidValue
    };
    odf_leaf(error, code)
}

#[cfg(test)]
mod tests {
    use super::native;
    use crate::error::{Code, Source};

    #[test]
    fn native_preserves_refusals_and_conflicts() {
        let refusal = native(litchi_odf_common::Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedTarget,
            "add_feature",
            "the target does not accept this feature",
        ));
        assert_eq!(refusal.code(), Code::UnsupportedType);
        assert!(matches!(refusal.leaf(), Source::Odf(_)));

        let conflict = native(litchi_odf_common::Error::patch_conflict(
            "add_feature",
            "the package base differs",
        ));
        assert_eq!(conflict.code(), Code::InvalidInput);
        assert!(matches!(conflict.leaf(), Source::Odf(_)));

        let unsupported = native(litchi_odf_common::Error::Core(
            litchi_core::Error::Unsupported("feature is unavailable".to_owned()),
        ));
        assert_eq!(unsupported.code(), Code::UnsupportedCapability);
        assert!(matches!(unsupported.leaf(), Source::Core(_)));
    }
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
            Some(Scalar::Text(value) | Scalar::Token(value)) => Self::Text(value.to_owned()),
            Some(Scalar::Bool(value)) => Self::Bool(value),
            Some(Scalar::Int(value)) => Self::Int(value),
            Some(Scalar::Number(value)) => Self::Number(value),
            Some(Scalar::Length(value)) => Self::Length(value),
            Some(Scalar::Color(value)) => Self::Color(value),
            Some(Scalar::List(items)) => Self::Text(items.join(", ")),
        }
    }

    fn text(&self) -> Option<String> {
        match self {
            Self::Null => None,
            Self::Text(value) => Some(value.clone()),
            Self::Bool(value) => Some(value.to_string()),
            Self::Int(value) => Some(value.to_string()),
            Self::Number(value) => Some(crate::value::DotNetDouble(*value).to_string()),
            Self::Length(value) => Some(format!("{}emu", value.emu())),
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
}

/// Admit `props` through `table` for `operation`: each property once.
fn admit<S: Copy + PartialEq, const N: usize, const A: usize>(
    table: &Table<S, N, A>,
    operation: Operation,
    props: &Props<'_>,
) -> Result<Vec<(S, &'static str, Arg)>> {
    let mut admitted: Vec<(S, &'static str, Arg)> = Vec::with_capacity(props.len());
    for (key, value) in props.iter() {
        let (_, entry) = table.resolve(key.as_str()).map_err(|error| {
            error
                .with_format(FileFormat::Ods)
                .with_property(key.as_str().to_owned())
        })?;
        let name = entry.spec.name;
        let (setter, arg) = entry
            .prepare(operation, value, |setter, scalar| {
                Ok((setter, Arg::of(scalar)))
            })
            .map_err(|error| error.with_format(FileFormat::Ods))?;
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

fn feature_path(sheet: u32, step: &str, index: Option<usize>) -> Path<'static> {
    let mut last = Step::new(step.to_owned());
    if let Some(index) = index {
        last = last.with(Predicate::Index(litchi_core::Position::new(index)));
    }
    Path::from_steps([
        Step::new("sheet").with(Predicate::Index(litchi_core::Position::new(sheet as usize))),
        last,
    ])
}

/// The highest zero-based row and column of an ODS grid.
const MAX_ROW: u32 = 1_048_575;
const MAX_COLUMN: u32 = 16_383;

/// Parse an A1 cell (dollar signs ignored) as zero-based `(row, column)`.
fn parse_cell(text: &str) -> Option<(u32, u32)> {
    let cell = litchi_core::selector::CellRef::parse(&text.trim().replace('$', ""))?;
    (cell.row() <= MAX_ROW && cell.column() <= MAX_COLUMN).then(|| (cell.row(), cell.column()))
}

/// Parse an A1 cell or range.
fn parse_rect(text: &str) -> Option<Rect> {
    let text = text.trim();
    match text.split_once(':') {
        Some((first, last)) => {
            let (first, last) = (parse_cell(first)?, parse_cell(last)?);
            Some((
                first.0.min(last.0),
                first.1.min(last.1),
                first.0.max(last.0),
                first.1.max(last.1),
            ))
        },
        None => parse_cell(text).map(|(row, column)| (row, column, row, column)),
    }
}

/// Space-separated A1 ranges.
fn parse_ranges(key: &'static str, value: &str) -> Result<Vec<Rect>> {
    let ranges = value
        .split_whitespace()
        .map(|range| parse_rect(range).ok_or_else(|| invalid(key, "the range is not an A1 range")))
        .collect::<Result<Vec<_>>>()?;
    if ranges.is_empty() {
        return Err(invalid(key, "the range is empty"));
    }
    Ok(ranges)
}

/// Where a feature is added: its worksheet, and the cell or range the
/// parent path named.
pub(super) struct Parent {
    pub(super) sheet: u32,
    pub(super) reference: Option<Rect>,
}

fn sheet_name(file: &Ods, sheet: u32) -> Result<String> {
    Ok(file
        .worksheets()?
        .get(sheet as usize)
        .ok_or_else(super::node::missing_sheet)?
        .name
        .clone())
}

/// Stage one sheet-feature edit and publish it.
fn publish(
    file: &Ods,
    limits: &SemanticLimits,
    stage: impl FnOnce(&mut litchi_ods::document::Edit) -> Result<()>,
) -> Result<Commit> {
    let mut edit = file.leaf_edit()?;
    stage(&mut edit)?;
    finish(file, edit, 1, limits)
}

// ---------------------------------------------------------------------------
// Dispatch

/// Add a worksheet feature under `parent`; `implied` is the rule type a
/// CF element word (`databar`, `cellis`) names.
pub(super) fn add(
    file: &Ods,
    parent: &Parent,
    kind: Kind,
    implied: Option<&'static str>,
    node: &NewNode<'_>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    if !matches!(node.placement(), None | Some(NodePlacement::End)) {
        return Err(error(
            Code::UnsupportedCapability,
            "worksheet features are added at the end of their kind",
        )
        .with_suggestion("omit placement or use Placement::End"));
    }
    let props = node.props();
    match kind {
        Kind::Hyperlink => add_hyperlink(file, parent, props, limits),
        Kind::Picture => add_drawing(file, parent, true, props, limits),
        Kind::Shape => add_drawing(file, parent, false, props, limits),
        Kind::Cf => add_cf(file, parent, implied, props, limits),
        Kind::DataValidation => add_validation(file, parent, props, limits),
        Kind::Autofilter => add_autofilter(file, parent, props, limits),
        Kind::RowBreak | Kind::ColBreak => add_break(file, parent, kind, props, limits),
        _ => Err(error(
            Code::UnsupportedElement,
            "this element kind cannot be added to an ODS worksheet",
        )),
    }
}

/// The feature a path names; `None` for other targets.
fn feature_target<'f>(file: &'f Ods, target: &TargetArg<'_>) -> Result<Option<Node<'f>>> {
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
    let node = crate::path::resolve::get(&super::Nav::new(file), path)?;
    Ok(matches!(node.item(), Item::Feature(_)).then_some(node))
}

/// A `set` on a feature path; `None` for other targets.
pub(super) fn try_set(
    file: &Ods,
    target: &TargetArg<'_>,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Option<Commit>> {
    let Some(node) = feature_target(file, target)? else {
        return Ok(None);
    };
    let Item::Feature(feature) = node.item() else {
        return Ok(None);
    };
    let sheet = node.sheet_position();
    Ok(Some(match feature.kind {
        Kind::Hyperlink => set_hyperlink(file, sheet, feature, props, limits)?,
        Kind::Picture | Kind::Shape => set_drawing(file, sheet, feature, props, limits)?,
        Kind::Cf => set_cf(file, sheet, feature, props, limits)?,
        Kind::DataValidation => set_validation(file, sheet, feature, props, limits)?,
        Kind::Autofilter => set_autofilter(file, sheet, feature, props, limits)?,
        _ => set_break(file, sheet, feature, props, limits)?,
    }))
}

/// Remove the feature a path names; `None` for other targets.
pub(super) fn try_remove(
    file: &Ods,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Option<Commit>> {
    let Some(node) = feature_target(file, target)? else {
        return Ok(None);
    };
    let Item::Feature(feature) = node.item() else {
        return Ok(None);
    };
    let sheet = node.sheet_position();
    let name = sheet_name(file, sheet)?;
    let index = feature.index as usize;
    Ok(Some(match &feature.data {
        Data::Hyperlink(link) => {
            let mut cell = existing_cell(file, sheet, link.row, link.column)?;
            cell.remove_hyperlink(link.slot)
                .ok_or_else(|| features::missing("the hyperlink is missing"))?;
            stage_cell(file, sheet, (link.row, link.column), cell, limits)?
        },
        Data::Drawing(_) => {
            let kind = drawing_kind(feature.kind);
            publish(file, limits, |edit| {
                edit.remove_sheet_drawing(&name, kind, index)
                    .map_err(native)
            })?
        },
        Data::Cf(_) => publish(file, limits, |edit| {
            edit.remove_sheet_format(&name, index).map_err(native)
        })?,
        Data::Validation(_) => publish(file, limits, |edit| {
            edit.remove_sheet_validation(&name, index).map_err(native)
        })?,
        Data::Autofilter(_) => publish(file, limits, |edit| {
            edit.set_sheet_auto_filter(&name, None).map_err(native)
        })?,
        Data::Break(_) => {
            let mut breaks = breaks(file, sheet, feature.kind)?;
            if index >= breaks.len() {
                return Err(features::missing("the page break is missing"));
            }
            breaks.remove(index);
            publish_breaks(file, sheet, feature.kind, &breaks, limits)?
        },
    }))
}

fn drawing_kind(kind: Kind) -> DrawingKind {
    if kind == Kind::Picture {
        DrawingKind::Picture
    } else {
        DrawingKind::Shape
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

/// The `xlink:href` of an in-document location (`Sheet2!A1` ->
/// `#Sheet2.A1`; a bare cell names `sheet`).
fn location_href(sheet: &str, location: &str) -> String {
    let location = location.trim().trim_start_matches('#');
    match location.rsplit_once('!') {
        Some((owner, cell)) => format!("#{owner}.{cell}"),
        None if parse_rect(location).is_some() => {
            let owner = if sheet
                .chars()
                .all(|char| char.is_alphanumeric() || char == '_')
            {
                sheet.to_owned()
            } else {
                format!("'{}'", sheet.replace('\'', "''"))
            };
            format!("#{owner}.{location}")
        },
        None => format!("#{location}"),
    }
}

/// The stored cell at a coordinate, or an empty one.
fn existing_cell(file: &Ods, sheet: u32, row: u32, column: u32) -> Result<Cell> {
    let data = file
        .worksheets()?
        .get(sheet as usize)
        .ok_or_else(super::node::missing_sheet)?;
    let mut cell = data
        .cell(row as usize, column as usize)
        .cloned()
        .unwrap_or_else(Cell::empty);
    cell.repeat = std::num::NonZeroUsize::MIN;
    Ok(cell)
}

fn stage_cell(
    file: &Ods,
    sheet: u32,
    at: (u32, u32),
    cell: Cell,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let mut changes = BTreeMap::new();
    changes.insert(sheet, BTreeMap::from([(at, cell)]));
    publish(file, limits, |edit| stage_cells(edit, changes))
}

/// A hyperlink request applied to a cell: the link spans the cell's whole
/// text, which `display` replaces.
struct LinkRequest {
    at: (u32, u32),
    href: Option<String>,
    display: Option<String>,
    tooltip: Option<Option<String>>,
}

fn link_cell(mut cell: Cell, slot: Option<usize>, request: &LinkRequest) -> Result<Cell> {
    let current = slot.and_then(|slot| cell.hyperlinks().get(slot).cloned());
    let href = request
        .href
        .clone()
        .or_else(|| current.as_ref().map(|link| link.href.clone()))
        .ok_or_else(|| missing("url", "a hyperlink needs a url or a location"))?;
    if let Some(display) = &request.display {
        if cell.formula.is_some() || !matches!(cell.value, CellValue::Text(_) | CellValue::Empty) {
            return Err(error(
                Code::UnsupportedCapability,
                "the display text of an ODS hyperlink is its cell's text; the cell holds a formula or a non-text value",
            )
            .with_property("display"));
        }
        cell.clear_hyperlinks();
        cell.value = CellValue::Text(display.clone());
        cell.text.clone_from(display);
    } else if cell.text.is_empty() {
        // An empty cell shows its target, as Calc does.
        let shown = href.trim_start_matches('#').to_owned();
        cell.value = CellValue::Text(shown.clone());
        cell.text = shown;
    }
    let title = match &request.tooltip {
        Some(title) => title.clone(),
        None => current.as_ref().and_then(|link| link.title.clone()),
    };
    // One link over the whole text.
    cell.clear_hyperlinks();
    let mut link = LeafLink::new(href);
    link.text.clone_from(&cell.text);
    link.title = title;
    let range = 0..cell.text.len();
    cell.add_hyperlink(range, link)
        .map_err(|error| leaf(error, Code::InvalidValue))?;
    Ok(cell)
}

fn link_request(
    file: &Ods,
    sheet: u32,
    operation: Operation,
    props: &Props<'_>,
    mut at: Option<(u32, u32)>,
) -> Result<LinkRequest> {
    let name = sheet_name(file, sheet)?;
    let mut href = None;
    let mut display = None;
    let mut tooltip = None;
    for (key, key_name, arg) in admit(&features::HYPERLINK_TABLE, operation, props)? {
        match key {
            HyperlinkKey::Ref => {
                let value = arg.text().unwrap_or_default();
                let rect = parse_rect(&value)
                    .ok_or_else(|| invalid(key_name, "the reference is not an A1 cell"))?;
                if rect.0 != rect.2 || rect.1 != rect.3 {
                    return Err(error(
                        Code::UnsupportedCapability,
                        "an ODS hyperlink is a text:a anchor in one cell",
                    )
                    .with_property(key_name)
                    .with_suggestion("add one hyperlink per cell"));
                }
                at = Some((rect.0, rect.1));
            },
            HyperlinkKey::Url => match arg.text() {
                Some(value) if value.starts_with('#') => {
                    href = Some(location_href(&name, &value));
                },
                Some(value)
                    if !value.contains("://") && !value.contains(':') && value.contains('!') =>
                {
                    href = Some(location_href(&name, &value));
                },
                Some(value) => {
                    check_url(key_name, &value)?;
                    href = Some(value);
                },
                None => {},
            },
            HyperlinkKey::Location => {
                if let Some(value) = arg.text() {
                    href = Some(location_href(&name, &value));
                }
            },
            HyperlinkKey::Display => display = arg.text(),
            HyperlinkKey::Tooltip => tooltip = Some(arg.text()),
        }
    }
    let at = at.ok_or_else(|| missing("ref", "a hyperlink needs a ref"))?;
    Ok(LinkRequest {
        at,
        href,
        display,
        tooltip,
    })
}

fn add_hyperlink(
    file: &Ods,
    parent: &Parent,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    let at = match parent.reference {
        Some((row, column, last_row, last_column)) if (row, column) == (last_row, last_column) => {
            Some((row, column))
        },
        _ => None,
    };
    let request = link_request(file, parent.sheet, Operation::Add, props, at)?;
    if request.href.is_none() {
        return Err(missing("url", "a hyperlink needs a url or a location"));
    }
    let cell = existing_cell(file, parent.sheet, request.at.0, request.at.1)?;
    // As XLSX's `put_hyperlink`: the cell's link is replaced.
    let cell = link_cell(cell, Some(0), &request)?;
    let commit = stage_cell(file, parent.sheet, request.at, cell, limits)?;
    let index = features::links(&commit.after, parent.sheet)?
        .iter()
        .position(|link| (link.row, link.column) == request.at)
        .unwrap_or_default();
    Ok((commit, feature_path(parent.sheet, "hyperlink", Some(index))))
}

fn set_hyperlink(
    file: &Ods,
    sheet: u32,
    feature: &Feature,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let Data::Hyperlink(link) = &feature.data else {
        return Err(error(Code::InternalError, "a hyperlink node holds no link"));
    };
    let request = link_request(
        file,
        sheet,
        Operation::Set,
        props,
        Some((link.row, link.column)),
    )?;
    let cell = existing_cell(file, sheet, link.row, link.column)?;
    if request.at == (link.row, link.column) {
        let cell = link_cell(cell, Some(link.slot), &request)?;
        return stage_cell(file, sheet, request.at, cell, limits);
    }
    // A moved link: the old cell keeps its text, the new one takes the link.
    let mut old = cell;
    let held = old
        .remove_hyperlink(link.slot)
        .ok_or_else(|| features::missing("the hyperlink is missing"))?;
    let request = LinkRequest {
        href: request.href.or(Some(held.href)),
        tooltip: request.tooltip.or(Some(held.title)),
        ..request
    };
    let target = existing_cell(file, sheet, request.at.0, request.at.1)?;
    let target = link_cell(target, Some(0), &request)?;
    let mut changes = BTreeMap::new();
    changes.insert(
        sheet,
        BTreeMap::from([((link.row, link.column), old), (request.at, target)]),
    );
    publish(file, limits, |edit| stage_cells(edit, changes))
}

// ---------------------------------------------------------------------------
// Pictures and shapes

/// The largest image a picture embeds.
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

/// The extension and media type of image bytes.
fn image_type(data: &[u8]) -> Option<(&'static str, &'static str)> {
    if data.starts_with(b"\x89PNG\r\n\x1a\n") {
        Some(("png", "image/png"))
    } else if data.starts_with(&[0xff, 0xd8, 0xff]) {
        Some(("jpg", "image/jpeg"))
    } else if data.starts_with(b"GIF87a") || data.starts_with(b"GIF89a") {
        Some(("gif", "image/gif"))
    } else {
        None
    }
}

/// A stable package name for image bytes (FNV-1a).
fn image_path(data: &[u8], extension: &str) -> String {
    let mut hash: u64 = 0xcbf2_9ce4_8422_2325;
    for byte in data {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x0100_0000_01b3);
    }
    format!("Pictures/litchi-{hash:016x}.{extension}")
}

/// A drawing length in centimetres: a length with a unit, or EMU.
fn length_cm(key: &'static str, arg: &Arg) -> Result<f64> {
    let emu = match arg {
        Arg::Length(length) => length.emu(),
        Arg::Text(text) if !text.trim().bytes().all(|byte| byte.is_ascii_digit()) => {
            Length::parse(text.trim())
                .map(Length::emu)
                .map_err(|_| invalid(key, "the value is a length with a unit"))?
        },
        _ => {
            return Err(invalid(
                key,
                "an ODS drawing is placed in lengths: give a length with a unit",
            )
            .with_suggestion("use a length such as 2cm or 914400emu"));
        },
    };
    #[allow(clippy::cast_precision_loss, reason = "drawing lengths are small")]
    let cm = emu as f64 / 360_000.0;
    Ok(cm)
}

/// A plain RGB color as `#rrggbb`.
fn rgb(key: &'static str, color: &Color) -> Result<String> {
    match color.base() {
        crate::value::Base::Rgb(value) if color.transforms().is_empty() => Ok(format!(
            "#{:02x}{:02x}{:02x}",
            value.red(),
            value.green(),
            value.blue()
        )),
        _ => Err(invalid(key, "the color must be an RGB color")),
    }
}

fn arg_color(key: &'static str, arg: &Arg, context: ColorContext) -> Result<String> {
    match arg {
        Arg::Color(color) => rgb(key, color),
        Arg::Text(value) => rgb(
            key,
            &Color::parse(value.trim(), context)
                .map_err(|_| invalid(key, "the value is not a color"))?,
        ),
        _ => Err(invalid(key, "the value is not a color")),
    }
}

fn shape_geometry(key: &'static str, value: &str) -> Result<&'static str> {
    let value = value.trim();
    if value.eq_ignore_ascii_case("rectangle") || value.eq_ignore_ascii_case("textbox") {
        return Ok("rectangle");
    }
    if value.eq_ignore_ascii_case("oval") {
        return Ok("ellipse");
    }
    GEOMETRIES
        .iter()
        .find(|(name, _)| name.eq_ignore_ascii_case(value))
        .map(|(_, odf)| *odf)
        .ok_or_else(|| {
            invalid(key, "the ODS shape editor writes the basic presets")
                .with_suggestion("use rect, roundRect, ellipse, diamond, triangle, rtTriangle, parallelogram, trapezoid, hexagon, octagon, star5 or rightArrow")
        })
}

/// A stroke value: `none`, a color, or `color:points`.
fn stroke(key: &'static str, value: &str) -> Result<Option<(String, String)>> {
    if value.trim().eq_ignore_ascii_case("none") {
        return Ok(None);
    }
    let mut parts = value.split(':');
    let color = rgb(
        key,
        &Color::parse(parts.next().unwrap_or("").trim(), ColorContext::Drawing)
            .map_err(|_| invalid(key, "the value is not a color"))?,
    )?;
    let width = match parts.next() {
        Some(width) => {
            let points: f64 = width
                .trim()
                .trim_end_matches("pt")
                .parse()
                .ok()
                .filter(|points: &f64| points.is_finite() && *points >= 0.0 && *points <= 1_584.0)
                .ok_or_else(|| invalid(key, "the line width is a number of points"))?;
            format!("{}pt", crate::value::DotNetDouble(points))
        },
        None => "0.75pt".to_owned(),
    };
    if parts
        .next()
        .is_some_and(|dash| !dash.trim().eq_ignore_ascii_case("solid"))
    {
        return Err(error(
            Code::UnsupportedProperty,
            "the ODS shape editor writes solid strokes",
        )
        .with_property(key));
    }
    Ok(Some((color, width)))
}

/// Drawing geometry in centimetres: `[x, y, width, height]`.
#[derive(Default)]
struct Frame([Option<f64>; 4]);

impl Frame {
    fn set(&mut self, axis: usize, key: &'static str, arg: &Arg) -> Result<()> {
        self.0[axis] = Some(length_cm(key, arg)?);
        Ok(())
    }

    fn is_empty(&self) -> bool {
        self.0.iter().all(Option::is_none)
    }

    fn resolve(&self, current: [f64; 4]) -> [f64; 4] {
        let mut out = current;
        for (slot, value) in out.iter_mut().zip(self.0) {
            if let Some(value) = value {
                *slot = value;
            }
        }
        out
    }
}

fn add_drawing(
    file: &Ods,
    parent: &Parent,
    picture: bool,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    let kind = if picture { Kind::Picture } else { Kind::Shape };
    let mut frame = Frame::default();
    let mut spec = DrawingSpec {
        name: String::new(),
        description: None,
        title: None,
        frame: [0.0; 4],
        end_cell: None,
        geometry: (!picture).then(|| "rectangle".to_owned()),
        text: None,
        fill: None,
        line: None,
        image: None,
    };
    let mut src = None;
    let mut textbox = false;
    if picture {
        for (key, name, arg) in admit(&features::PICTURE_TABLE, Operation::Add, props)? {
            match key {
                PictureKey::Src => src = arg.text(),
                PictureKey::Name => spec.name = arg.text().unwrap_or_default(),
                PictureKey::Alt => spec.description = arg.text(),
                PictureKey::Title => spec.title = arg.text(),
                PictureKey::X => frame.set(0, name, &arg)?,
                PictureKey::Y => frame.set(1, name, &arg)?,
                PictureKey::Width => frame.set(2, name, &arg)?,
                PictureKey::Height => frame.set(3, name, &arg)?,
                _ => {},
            }
        }
    } else {
        for (key, name, arg) in admit(&features::SHAPE_TABLE, Operation::Add, props)? {
            let text = || arg.text().unwrap_or_default();
            match key {
                ShapeKey::Name => spec.name = text(),
                ShapeKey::Text => spec.text = Some(text().replace("\\n", "\n")),
                ShapeKey::Geometry => {
                    textbox = text().trim().eq_ignore_ascii_case("textbox");
                    spec.geometry = Some(shape_geometry(name, &text())?.to_owned());
                },
                ShapeKey::X => frame.set(0, name, &arg)?,
                ShapeKey::Y => frame.set(1, name, &arg)?,
                ShapeKey::Width => frame.set(2, name, &arg)?,
                ShapeKey::Height => frame.set(3, name, &arg)?,
                ShapeKey::Fill => {
                    let value = text();
                    spec.fill = Some(if value.trim().eq_ignore_ascii_case("none") {
                        None
                    } else {
                        Some(arg_color(name, &arg, ColorContext::Drawing)?)
                    });
                },
                ShapeKey::Line => spec.line = Some(stroke(name, &text())?),
                ShapeKey::Alt => spec.description = arg.text(),
                ShapeKey::Hyperlink | ShapeKey::Pending => {},
            }
        }
        if textbox {
            spec.fill.get_or_insert(None);
            spec.line.get_or_insert(None);
        }
    }
    // Defaults: a 5.08 cm square for a picture, a 5 by 2.5 cm shape, at
    // the parent cell's corner when the parent is a cell (approximated by
    // LibreOffice's default 2.258 cm columns and 0.452 cm rows).
    let origin = parent.reference.map_or([0.0, 0.0], |(row, column, _, _)| {
        [f64::from(column) * 2.258, f64::from(row) * 0.452]
    });
    let defaults = if picture {
        [origin[0], origin[1], 5.08, 5.08]
    } else {
        [origin[0], origin[1], 5.0, 2.5]
    };
    spec.frame = frame.resolve(defaults);
    let existing = features::list(file, parent.sheet, kind)?;
    if spec.name.is_empty() {
        let prefix = if picture { "Picture" } else { "Shape" };
        spec.name = format!("{prefix} {}", existing.len() + 1);
    }
    let resource = if picture {
        let src = src.ok_or_else(|| missing("src", "a picture needs a src"))?;
        let data = read_source(&src)?;
        let (extension, media_type) = image_type(&data).ok_or_else(|| {
            error(
                Code::UnsupportedCapability,
                "the ODS picture editor embeds PNG, JPEG and GIF images",
            )
            .with_property("src")
        })?;
        let path = image_path(&data, extension);
        spec.image = Some(path.clone());
        Some((
            Resource::new(path, media_type, data).map_err(native)?,
            Collision::ReuseEquivalent,
        ))
    } else {
        None
    };
    let name = sheet_name(file, parent.sheet)?;
    let mut index = 0;
    let commit = publish(file, limits, |edit| {
        index = edit
            .add_sheet_drawing(&name, &spec, resource)
            .map_err(native)?;
        Ok(())
    })?;
    Ok((
        commit,
        feature_path(parent.sheet, features::step_name(kind), Some(index)),
    ))
}

fn set_drawing(
    file: &Ods,
    sheet: u32,
    feature: &Feature,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let Data::Drawing(drawing) = &feature.data else {
        return Err(error(
            Code::InternalError,
            "a drawing node holds no drawing",
        ));
    };
    let picture = feature.kind == Kind::Picture;
    let mut change = DrawingChange::default();
    let mut frame = Frame::default();
    let mut set_name = |value: Option<String>| -> Result<()> {
        change.name = Some(
            value
                .filter(|name| !name.is_empty())
                .ok_or_else(|| invalid("name", "a drawing name cannot be empty"))?,
        );
        Ok(())
    };
    let mut description = None;
    if picture {
        for (key, name, arg) in admit(&features::PICTURE_TABLE, Operation::Set, props)? {
            match key {
                PictureKey::Name => set_name(arg.text())?,
                PictureKey::Alt => description = Some(arg.text()),
                PictureKey::X => frame.set(0, name, &arg)?,
                PictureKey::Y => frame.set(1, name, &arg)?,
                PictureKey::Width => frame.set(2, name, &arg)?,
                PictureKey::Height => frame.set(3, name, &arg)?,
                _ => {
                    return Err(
                        error(Code::UnsupportedProperty, "the property is not settable")
                            .with_property(name),
                    );
                },
            }
        }
    } else {
        for (key, name, arg) in admit(&features::SHAPE_TABLE, Operation::Set, props)? {
            match key {
                ShapeKey::Name => set_name(arg.text())?,
                ShapeKey::Alt => description = Some(arg.text()),
                ShapeKey::X => frame.set(0, name, &arg)?,
                ShapeKey::Y => frame.set(1, name, &arg)?,
                ShapeKey::Width => frame.set(2, name, &arg)?,
                ShapeKey::Height => frame.set(3, name, &arg)?,
                _ => {
                    return Err(
                        error(Code::UnsupportedProperty, "the property is not settable")
                            .with_property(name),
                    );
                },
            }
        }
    }
    change.description = description;
    if !frame.is_empty() {
        if drawing.cell.is_some() && (frame.0[0].is_some() || frame.0[1].is_some()) {
            return Err(error(
                Code::UnsupportedCapability,
                "a cell-anchored ODS drawing is positioned by its cell; its x and y are offsets the unified layer does not author",
            )
            .with_property("x"));
        }
        let current = [
            drawing.frame[0].unwrap_or(0.0),
            drawing.frame[1].unwrap_or(0.0),
            drawing.frame[2].unwrap_or(1.0),
            drawing.frame[3].unwrap_or(1.0),
        ];
        change.frame = Some(frame.resolve(current));
    }
    let name = sheet_name(file, sheet)?;
    let kind = drawing_kind(feature.kind);
    let index = feature.index as usize;
    publish(file, limits, |edit| {
        edit.set_sheet_drawing(&name, kind, index, &change)
            .map_err(native)
    })
}

// ---------------------------------------------------------------------------
// Conditional formats

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

/// The CF element words OfficeCLI adds directly (`--type databar`).
pub(super) fn cf_element(name: &str) -> Option<&'static str> {
    Some(match name.to_ascii_lowercase().as_str() {
        "cf" | "conditionalformatting" | "cfextended" => "",
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
        _ => return None,
    })
}

fn cf_operator(key: &'static str, value: &str) -> Result<usize> {
    Ok(match value.trim().to_ascii_lowercase().as_str() {
        "between" => 0,
        "notbetween" => 1,
        "equal" | "eq" | "=" | "==" => 2,
        "notequal" | "ne" | "!=" | "<>" => 3,
        "greaterthan" | "gt" | ">" => 4,
        "lessthan" | "lt" | "<" => 5,
        "greaterthanorequal" | "gte" | "ge" | ">=" => 6,
        "lessthanorequal" | "lte" | "le" | "<=" => 7,
        _ => return Err(invalid(key, "the operator is not a cellIs comparison")),
    })
}

/// A threshold: its entry type and value (OpenFormula for a formula).
type Threshold = (EntryType, String);

fn cf_threshold(key: &'static str, value: Option<&str>, fallback: EntryType) -> Result<Threshold> {
    let value = value
        .map(str::trim)
        .filter(|value| !value.is_empty() && !value.eq_ignore_ascii_case("auto"));
    let Some(value) = value else {
        return Ok((fallback, "0".to_owned()));
    };
    let (kind, text) = if let Some(rest) = value.strip_suffix('%') {
        (EntryType::Percent, rest)
    } else if let Some((prefix, rest)) = value.split_once(':') {
        let kind = match prefix.to_ascii_lowercase().as_str() {
            "percent" => EntryType::Percent,
            "percentile" => EntryType::Percentile,
            "num" | "number" => EntryType::Number,
            "formula" => EntryType::Formula,
            "min" => EntryType::Minimum,
            "max" => EntryType::Maximum,
            _ => {
                return Err(invalid(
                    key,
                    "the threshold kind is num, percent, percentile or formula",
                ));
            },
        };
        (kind, rest)
    } else {
        (EntryType::Number, value)
    };
    if kind == EntryType::Formula {
        return Ok((kind, to_odff(text.trim())));
    }
    if text.trim().parse::<f64>().is_err() {
        return Err(invalid(key, "the threshold is not a number"));
    }
    Ok((kind, text.trim().to_owned()))
}

/// The parts of a CF request.
#[derive(Default)]
struct CfRequest {
    kind: Option<CfType>,
    ranges: Option<Vec<Rect>>,
    operator: Option<usize>,
    value: Option<String>,
    value2: Option<String>,
    formula: Option<String>,
    text: Option<String>,
    rank: Option<u32>,
    percent: Option<bool>,
    bottom: Option<bool>,
    above_average: Option<bool>,
    equal_average: Option<bool>,
    period: Option<usize>,
    color: Option<String>,
    negative_color: Option<Option<String>>,
    axis_color: Option<Option<String>>,
    axis_position: Option<Option<cf::DataBarAxisPosition>>,
    min: Option<String>,
    max: Option<String>,
    min_length: Option<u32>,
    max_length: Option<u32>,
    show_value: Option<bool>,
    min_color: Option<String>,
    mid_color: Option<String>,
    max_color: Option<String>,
    midpoint: Option<String>,
    iconset: Option<cf::IconSetType>,
    style: FormatStyle,
}

fn merge_style(target: &mut FormatStyle, change: &FormatStyle) {
    if change.background.is_some() {
        target.background.clone_from(&change.background);
    }
    if change.color.is_some() {
        target.color.clone_from(&change.color);
    }
    if change.bold.is_some() {
        target.bold = change.bold;
    }
    if change.italic.is_some() {
        target.italic = change.italic;
    }
    if change.underline.is_some() {
        target.underline = change.underline;
    }
    if change.strike.is_some() {
        target.strike = change.strike;
    }
    if change.size_pt.is_some() {
        target.size_pt = change.size_pt;
    }
    if change.font.is_some() {
        target.font.clone_from(&change.font);
    }
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
            CfKey::Period => {
                let value = text();
                request.period = Some(
                    PERIODS
                        .iter()
                        .position(|(period, _)| period.eq_ignore_ascii_case(value.trim()))
                        .ok_or_else(|| invalid(name, "the period is not a time period"))?,
                );
            },
            CfKey::Color => request.color = Some(color()?),
            CfKey::NegativeColor => request.negative_color = Some(Some(color()?)),
            CfKey::AxisColor => request.axis_color = Some(Some(color()?)),
            CfKey::AxisPosition => {
                request.axis_position = Some(match text().to_ascii_lowercase().as_str() {
                    "automatic" | "auto" => None,
                    "middle" => Some(cf::DataBarAxisPosition::Middle),
                    "none" => Some(cf::DataBarAxisPosition::None),
                    _ => return Err(invalid(name, "axisPosition is automatic, middle or none")),
                });
            },
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
                    cf::IconSetType::parse(value.trim())
                        .map_err(|_| invalid(name, "the icon set is not an ODF icon set"))?,
                );
            },
            CfKey::Fill => request.style.background = Some(color()?),
            CfKey::FontColor => request.style.color = Some(color()?),
            CfKey::FontBold => request.style.bold = Some(arg.flag()),
            CfKey::FontItalic => request.style.italic = Some(arg.flag()),
            CfKey::FontStrike => request.style.strike = Some(arg.flag()),
            CfKey::FontUnderline => {
                request.style.underline = match text().to_ascii_lowercase().as_str() {
                    "single" | "true" => Some(true),
                    "none" | "false" => Some(false),
                    "double" => {
                        return Err(error(
                            Code::UnsupportedProperty,
                            "the ODS conditional style writes single underlines",
                        )
                        .with_property(name));
                    },
                    _ => return Err(invalid(name, "underline is single or none")),
                };
            },
            CfKey::FontSize => {
                if let Arg::Length(length) = arg {
                    let points = length.points();
                    if !(1.0..=409.55).contains(&points) {
                        return Err(invalid(name, "the font size is 1 to 409.55 points"));
                    }
                    request.style.size_pt = Some((points * 100.0).round() / 100.0);
                }
            },
            CfKey::FontName => request.style.font = arg.text(),
            CfKey::Priority | CfKey::DxfId | CfKey::Pending => {},
        }
    }
    Ok(request)
}

/// The state of one format as the adapter edits it.
#[derive(Debug, Clone)]
struct CfState {
    kind: CfKind,
    operator: usize,
    values: Vec<String>,
    formula: Option<String>,
    text: Option<String>,
    rank: u32,
    percent: bool,
    bottom: bool,
    below: bool,
    equal: bool,
    period: usize,
    /// The first rule of a color scale, data bar or icon set, edited in
    /// place.
    graphic: Option<cf::Rule>,
}

impl CfState {
    fn fresh(kind: &CfType) -> Self {
        Self {
            kind: kind.kind,
            operator: 4,
            values: Vec::new(),
            formula: None,
            text: None,
            rank: 10,
            percent: kind.percent,
            bottom: kind.bottom,
            below: kind.below,
            equal: false,
            period: 0,
            graphic: None,
        }
    }

    fn of(feature: &features::CfFeature) -> Option<Self> {
        let rule = feature.rule.as_ref()?;
        Some(Self {
            kind: rule.kind,
            operator: rule.operator.unwrap_or(4),
            values: rule.values.clone(),
            formula: rule.formula.clone(),
            text: rule.text.clone(),
            rank: rule.rank.unwrap_or(10),
            percent: rule.percent,
            bottom: rule.bottom,
            below: rule.below,
            equal: rule.equal,
            period: rule.period.unwrap_or(0),
            graphic: matches!(
                rule.kind,
                CfKind::ColorScale | CfKind::DataBar | CfKind::IconSet
            )
            .then(|| feature.format.rules.first().cloned())
            .flatten(),
        })
    }

    fn takes_style(&self) -> bool {
        !matches!(
            self.kind,
            CfKind::ColorScale | CfKind::DataBar | CfKind::IconSet
        )
    }

    /// Apply a request; `fresh` fills the defaults a new rule needs.
    fn apply(&mut self, request: &CfRequest, fresh: bool) -> Result<()> {
        if let Some(operator) = request.operator {
            self.operator = operator;
        }
        if request.value.is_some() || request.value2.is_some() {
            let mut values = self.values.clone();
            values.resize(2, String::new());
            if let Some(value) = &request.value {
                values[0].clone_from(value);
            }
            if let Some(value) = &request.value2 {
                values[1].clone_from(value);
            }
            self.values = values
                .into_iter()
                .filter(|value| !value.is_empty())
                .collect();
        }
        if request.formula.is_some() {
            self.formula.clone_from(&request.formula);
        }
        if request.text.is_some() {
            self.text.clone_from(&request.text);
        }
        if let Some(rank) = request.rank {
            self.rank = rank;
        }
        if let Some(percent) = request.percent {
            self.percent = percent;
        }
        if let Some(bottom) = request.bottom {
            self.bottom = bottom;
        }
        if let Some(above) = request.above_average {
            self.below = !above;
        }
        if let Some(equal) = request.equal_average {
            self.equal = equal;
        }
        if let Some(period) = request.period {
            self.period = period;
        }
        match self.kind {
            CfKind::ColorScale => self.apply_scale(request, fresh),
            CfKind::DataBar => self.apply_bar(request, fresh),
            CfKind::IconSet => self.apply_icons(request, fresh),
            _ => Ok(()),
        }
    }

    fn apply_scale(&mut self, request: &CfRequest, fresh: bool) -> Result<()> {
        let mut entries = match (&self.graphic, fresh) {
            (Some(cf::Rule::ColorScale(scale)), false) => scale.entries.clone(),
            _ => vec![
                cf::ColorScaleEntry::new(EntryType::Minimum, "0", "#f8696b"),
                cf::ColorScaleEntry::new(EntryType::Maximum, "0", "#63be7b"),
            ],
        };
        if request.mid_color.is_some() || request.midpoint.is_some() {
            if entries.len() == 2 {
                entries.insert(
                    1,
                    cf::ColorScaleEntry::new(EntryType::Percentile, "50", "#ffeb84"),
                );
            }
        }
        if let Some(color) = &request.min_color {
            entries[0].color.clone_from(color);
        }
        if let Some(color) = &request.max_color {
            if let Some(last) = entries.last_mut() {
                last.color.clone_from(color);
            }
        }
        if entries.len() == 3 {
            if let Some(color) = &request.mid_color {
                entries[1].color.clone_from(color);
            }
            if request.midpoint.is_some() {
                let (kind, value) = cf_threshold(
                    "midpoint",
                    request.midpoint.as_deref(),
                    EntryType::Percentile,
                )?;
                entries[1].entry_type = kind;
                entries[1].value = value;
            }
        }
        if request.min.is_some() {
            let (kind, value) = cf_threshold("min", request.min.as_deref(), EntryType::Minimum)?;
            entries[0].entry_type = kind;
            entries[0].value = value;
        }
        if request.max.is_some() {
            let (kind, value) = cf_threshold("max", request.max.as_deref(), EntryType::Maximum)?;
            if let Some(last) = entries.last_mut() {
                last.entry_type = kind;
                last.value = value;
            }
        }
        self.graphic = Some(cf::Rule::ColorScale(cf::ColorScale::new(entries)));
        Ok(())
    }

    fn apply_bar(&mut self, request: &CfRequest, fresh: bool) -> Result<()> {
        let mut bar = match (&self.graphic, fresh) {
            (Some(cf::Rule::DataBar(bar)), false) => bar.clone(),
            _ => cf::DataBar::new(vec![
                cf::DataBarEntry::new(EntryType::AutomaticMinimum, "0"),
                cf::DataBarEntry::new(EntryType::AutomaticMaximum, "0"),
            ])
            .with_colors("#638ec6", None),
        };
        if bar.entries.len() != 2 {
            return Err(error(
                Code::UnsupportedCapability,
                "the data bar does not have two limits",
            ));
        }
        if let Some(color) = &request.color {
            bar.positive_color = Some(color.clone());
        }
        if let Some(color) = &request.negative_color {
            bar.negative_color.clone_from(color);
        }
        if let Some(color) = &request.axis_color {
            bar.axis_color.clone_from(color);
        }
        if let Some(position) = request.axis_position {
            bar.axis_position = position;
        }
        if request.min.is_some() {
            let (kind, value) =
                cf_threshold("min", request.min.as_deref(), EntryType::AutomaticMinimum)?;
            bar.entries[0] = cf::DataBarEntry::new(kind, value);
        }
        if request.max.is_some() {
            let (kind, value) =
                cf_threshold("max", request.max.as_deref(), EntryType::AutomaticMaximum)?;
            bar.entries[1] = cf::DataBarEntry::new(kind, value);
        }
        if let Some(length) = request.min_length {
            bar.min_length = Some(length.to_string());
        }
        if let Some(length) = request.max_length {
            bar.max_length = Some(length.to_string());
        }
        if let Some(show) = request.show_value {
            bar.show_value = Some(show);
        }
        self.graphic = Some(cf::Rule::DataBar(bar));
        Ok(())
    }

    fn apply_icons(&mut self, request: &CfRequest, fresh: bool) -> Result<()> {
        let mut icons = match (&self.graphic, fresh) {
            (Some(cf::Rule::IconSet(icons)), false) => icons.clone(),
            _ => icon_set(cf::IconSetType::ThreeTrafficLights1),
        };
        if let Some(kind) = request.iconset {
            let show = icons.show_value;
            icons = icon_set(kind);
            icons.show_value = show;
        }
        if let Some(show) = request.show_value {
            icons.show_value = Some(show);
        }
        self.graphic = Some(cf::Rule::IconSet(icons));
        Ok(())
    }

    /// The leaf format over `ranges` of `sheet`.
    fn format(&self, sheet: &str, ranges: &[Rect]) -> Result<cf::Format> {
        let addresses: Vec<String> = ranges.iter().map(|rect| address(sheet, *rect)).collect();
        let base = ranges
            .first()
            .map(|rect| address(sheet, (rect.0, rect.1, rect.0, rect.1)));
        let first = ranges.first().map_or_else(
            || "A1".to_owned(),
            |rect| {
                crate::format::workbook::Label::cell(rect.1, rect.0)
                    .as_str()
                    .to_owned()
            },
        );
        let odff = |key: &'static str, value: &str| -> Result<String> {
            let value = value.trim();
            if value.is_empty() {
                return Err(invalid(key, "the value cannot be empty"));
            }
            Ok(to_odff(value.strip_prefix('=').unwrap_or(value)))
        };
        let quoted = |text: &str| format!("\"{}\"", text.replace('"', "\"\""));
        let condition = |text: String| -> cf::Rule {
            let mut condition = cf::Condition::new(text, "Default");
            condition.base_cell_address.clone_from(&base);
            cf::Rule::Condition(condition)
        };
        let rule = match self.kind {
            CfKind::CellIs => {
                let (_, spelling) = CF_OPERATORS[self.operator];
                let first = odff(
                    "value",
                    self.values
                        .first()
                        .ok_or_else(|| missing("value", "a cellIs rule needs a value"))?,
                )?;
                if self.operator < 2 {
                    let second = odff(
                        "value2",
                        self.values
                            .get(1)
                            .ok_or_else(|| missing("value2", "a between rule needs value2"))?,
                    )?;
                    condition(format!("{spelling}({first},{second})"))
                } else {
                    condition(format!("{spelling}{first}"))
                }
            },
            CfKind::Formula => condition(format!(
                "formula-is({})",
                odff(
                    "formula",
                    self.formula
                        .as_deref()
                        .ok_or_else(|| missing("formula", "a formula rule needs a formula"))?
                )?
            )),
            CfKind::Text(operator) => {
                let text = self
                    .text
                    .as_deref()
                    .ok_or_else(|| missing("text", "a text rule needs its text"))?;
                let name = [
                    "contains-text",
                    "not-contains-text",
                    "begins-with",
                    "ends-with",
                ][usize::from(operator)];
                condition(format!("{name}({})", quoted(text)))
            },
            CfKind::ContainsBlanks => condition(format!(
                "formula-is({})",
                to_odff(&format!("LEN(TRIM({first}))=0"))
            )),
            CfKind::NotContainsBlanks => condition(format!(
                "formula-is({})",
                to_odff(&format!("LEN(TRIM({first}))>0"))
            )),
            CfKind::ContainsErrors => condition("is-error".to_owned()),
            CfKind::NotContainsErrors => condition("is-no-error".to_owned()),
            CfKind::DuplicateValues => condition("duplicate".to_owned()),
            CfKind::UniqueValues => condition("unique".to_owned()),
            CfKind::TopN => {
                let name = match (self.bottom, self.percent) {
                    (false, false) => "top-elements",
                    (true, false) => "bottom-elements",
                    (false, true) => "top-percent",
                    (true, true) => "bottom-percent",
                };
                condition(format!("{name}({})", self.rank))
            },
            CfKind::AboveAverage => condition(
                match (self.below, self.equal) {
                    (false, false) => "above-average",
                    (true, false) => "below-average",
                    (false, true) => "above-equal-average",
                    (true, true) => "below-equal-average",
                }
                .to_owned(),
            ),
            CfKind::TimePeriod => {
                cf::Rule::DateIs(cf::DateIs::new(PERIODS[self.period].1, "Default"))
            },
            CfKind::ColorScale | CfKind::DataBar | CfKind::IconSet => self
                .graphic
                .clone()
                .ok_or_else(|| error(Code::InternalError, "the graphic rule is missing"))?,
        };
        cf::Format::new(addresses, vec![rule]).map_err(|error| leaf(error, Code::InvalidValue))
    }
}

/// An icon set of `kind` with its default percent thresholds.
fn icon_set(kind: cf::IconSetType) -> cf::IconSet {
    let count = kind.as_str().bytes().next().map_or(3, |digit| digit - b'0');
    let entries = (0..count)
        .map(|index| {
            let value = u32::from(index) * 100 / u32::from(count);
            cf::IconSetEntry::new(EntryType::Percent, value.to_string())
        })
        .collect();
    cf::IconSet::new(kind, entries)
}

fn add_cf(
    file: &Ods,
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
        (None, Some(implied)) => {
            cf_type(implied).ok_or_else(|| error(Code::InternalError, "unknown implied type"))?
        },
        // OfficeCLI's `--type cf` without a type is a data bar.
        (None, None) => {
            cf_type("dataBar").ok_or_else(|| error(Code::InternalError, "no data bar"))?
        },
    };
    let ranges = match request.ranges.take() {
        Some(ranges) => ranges,
        None => vec![parent.reference.unwrap_or((0, 0, 9, 0))],
    };
    let mut state = CfState::fresh(&kind);
    state.apply(&request, true)?;
    if !request.style.is_empty() && !state.takes_style() {
        return Err(invalid(
            "fill",
            "a color scale, data bar or icon set takes no formatting",
        ));
    }
    let name = sheet_name(file, parent.sheet)?;
    let format = state.format(&name, &ranges)?;
    let mut index = 0;
    let commit = publish(file, limits, |edit| {
        index = edit
            .add_sheet_format(&name, &format, &request.style)
            .map_err(native)?;
        Ok(())
    })?;
    Ok((commit, feature_path(parent.sheet, "cf", Some(index))))
}

fn set_cf(
    file: &Ods,
    sheet: u32,
    feature: &Feature,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let Data::Cf(current) = &feature.data else {
        return Err(error(
            Code::InternalError,
            "a conditional-format node holds no format",
        ));
    };
    if current.format.rules.len() > 1 {
        return Err(error(
            Code::UnsupportedCapability,
            "the conditional format holds several rules; the unified layer edits single-rule formats",
        )
        .with_suggestion("remove the format and add one format per rule"));
    }
    let request = cf_request(Operation::Set, props)?;
    let mut state = CfState::of(current).ok_or_else(|| {
        error(
            Code::UnsupportedCapability,
            "the conditional-format rule is not one the unified layer authors",
        )
    })?;
    state.apply(&request, false)?;
    let mut style = current.style.clone();
    if !request.style.is_empty() {
        if !state.takes_style() {
            return Err(invalid(
                "fill",
                "a color scale, data bar or icon set takes no formatting",
            ));
        }
        merge_style(&mut style, &request.style);
    }
    let ranges = match &request.ranges {
        Some(ranges) => ranges.clone(),
        None => current
            .format
            .target_range_addresses
            .iter()
            .filter_map(|range| sf::parse_address(range).map(|(_, rect)| rect))
            .collect(),
    };
    let name = sheet_name(file, sheet)?;
    let format = state.format(&name, &ranges)?;
    let index = feature.index as usize;
    publish(file, limits, |edit| {
        edit.set_sheet_format(&name, index, &format, &style)
            .map_err(native)
    })
}

// ---------------------------------------------------------------------------
// Validations

/// The state of a validation as the adapter edits it.
#[derive(Debug, Clone)]
struct DvState {
    ranges: Vec<Rect>,
    kind: String,
    operator: usize,
    formula1: Option<String>,
    formula2: Option<String>,
    allow_blank: bool,
    dropdown: bool,
    show_input: bool,
    show_error: bool,
    error_style: String,
    prompt_title: Option<String>,
    prompt: Option<String>,
    error_title: Option<String>,
    error: Option<String>,
}

const VALIDATION_TYPES: [&str; 8] = [
    "none",
    "whole",
    "decimal",
    "list",
    "date",
    "time",
    "textLength",
    "custom",
];

impl DvState {
    fn of(validation: &sf::Validation) -> Self {
        let condition = features::parse_condition(&validation.condition);
        let (help_title, help, help_shown) = validation.help.clone().unwrap_or((None, None, false));
        let (style, error_title, error, error_shown) =
            validation
                .error
                .clone()
                .unwrap_or(("stop".to_owned(), None, None, false));
        Self {
            ranges: validation.ranges.clone(),
            kind: condition.as_ref().map_or_else(
                || {
                    if validation.condition.is_empty() {
                        "none".to_owned()
                    } else {
                        "custom".to_owned()
                    }
                },
                |condition| condition.kind.to_owned(),
            ),
            operator: condition
                .as_ref()
                .and_then(|condition| condition.operator)
                .unwrap_or(0),
            formula1: condition
                .as_ref()
                .and_then(|condition| condition.formula1.clone()),
            formula2: condition.and_then(|condition| condition.formula2),
            allow_blank: validation.allow_empty,
            dropdown: validation.display_list,
            show_input: help_shown,
            show_error: error_shown,
            error_style: style,
            prompt_title: help_title,
            prompt: help,
            error_title,
            error,
        }
    }

    fn apply(&mut self, props: &Props<'_>, operation: Operation) -> Result<()> {
        for (key, name, arg) in admit(&features::VALIDATION_TABLE, operation, props)? {
            let text = || arg.text().unwrap_or_default();
            match key {
                ValidationKey::Type => {
                    let value = text();
                    let value = if value.eq_ignore_ascii_case("any") {
                        "none".to_owned()
                    } else {
                        value
                    };
                    self.kind = VALIDATION_TYPES
                        .iter()
                        .find(|known| known.eq_ignore_ascii_case(value.trim()))
                        .map(|known| (*known).to_owned())
                        .ok_or_else(|| {
                            invalid(name, "the validation type is not one Excel knows")
                        })?;
                },
                ValidationKey::Ref => self.ranges = parse_ranges(name, &text())?,
                ValidationKey::Operator => {
                    let value = text();
                    self.operator = VALIDATION_OPERATORS
                        .iter()
                        .position(|(known, _)| known.eq_ignore_ascii_case(value.trim()))
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
                    self.error_style = match text().to_ascii_lowercase().as_str() {
                        "stop" => "stop",
                        "warning" | "warn" => "warning",
                        "information" | "info" => "information",
                        _ => {
                            return Err(invalid(
                                name,
                                "errorStyle is stop, warning or information",
                            ));
                        },
                    }
                    .to_owned();
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
        Ok(())
    }

    /// The `table:condition` of the state.
    fn condition(&self) -> Result<String> {
        let operand = |key: &'static str, value: &Option<String>| -> Result<String> {
            let value = value
                .as_deref()
                .ok_or_else(|| missing(key, "a typed validation needs formula1"))?
                .trim();
            Ok(to_odff(value.strip_prefix('=').unwrap_or(value)))
        };
        let comparison = |subject: &str| -> Result<String> {
            let (_, spelling) = VALIDATION_OPERATORS[self.operator];
            let first = operand("formula1", &self.formula1)?;
            if self.operator < 2 {
                let second = operand("formula2", &self.formula2)
                    .map_err(|_| missing("formula2", "a between validation needs formula2"))?;
                Ok(format!("{subject}-is-{spelling}({first},{second})"))
            } else {
                Ok(format!("{subject}(){spelling}{first}"))
            }
        };
        Ok(match self.kind.as_str() {
            "none" => String::new(),
            "list" => {
                let value = self
                    .formula1
                    .as_deref()
                    .ok_or_else(|| missing("formula1", "a list validation needs formula1"))?
                    .trim();
                if let Some(formula) = value.strip_prefix('=') {
                    format!("of:cell-content-is-in-list({})", to_odff(formula))
                } else {
                    let literal = value
                        .strip_prefix('"')
                        .and_then(|rest| rest.strip_suffix('"'))
                        .unwrap_or(value);
                    let reference = parse_rect(literal.trim_start_matches('$')).is_some()
                        || literal.contains('!');
                    if reference && !value.starts_with('"') && !literal.contains(',') {
                        format!("of:cell-content-is-in-list({})", to_odff(literal))
                    } else {
                        let items: Vec<String> = literal
                            .split(',')
                            .map(|item| format!("\"{}\"", item.replace('"', "\"\"")))
                            .collect();
                        format!("of:cell-content-is-in-list({})", items.join(";"))
                    }
                }
            },
            "custom" => format!(
                "of:is-true-formula({})",
                operand("formula1", &self.formula1)?
            ),
            "textLength" => format!("of:{}", comparison("cell-content-text-length")?),
            kind => {
                let head = match kind {
                    "whole" => "cell-content-is-whole-number()",
                    "decimal" => "cell-content-is-decimal-number()",
                    "date" => "cell-content-is-date()",
                    _ => "cell-content-is-time()",
                };
                let tail = comparison("cell-content")?;
                // `cell-content-is-between(…)`, `cell-content()>=…`.
                format!("of:{head} and {tail}")
            },
        })
    }

    fn leaf(&self) -> Result<ValidationSpec> {
        Ok(ValidationSpec {
            ranges: self.ranges.clone(),
            condition: self.condition()?,
            allow_empty: self.allow_blank,
            display_list: self.dropdown,
            help: (
                self.prompt_title.clone(),
                self.prompt.clone(),
                self.show_input,
            ),
            error: (
                self.error_style.clone(),
                self.error_title.clone(),
                self.error.clone(),
                self.show_error,
            ),
        })
    }
}

fn overlaps(left: &[Rect], right: &[Rect]) -> bool {
    left.iter().any(|a| {
        right
            .iter()
            .any(|b| a.0 <= b.2 && b.0 <= a.2 && a.1 <= b.3 && b.1 <= a.3)
    })
}

fn add_validation(
    file: &Ods,
    parent: &Parent,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    let named = props.iter().any(|(key, _)| {
        ["ref", "sqref", "range"]
            .iter()
            .any(|name| key.as_str().eq_ignore_ascii_case(name))
    });
    let ranges = match (parent.reference, named) {
        (_, true) => vec![(0, 0, 0, 0)],
        (Some(rect), false) => vec![rect],
        (None, false) => return Err(missing("ref", "a data validation needs a ref")),
    };
    let mut state = DvState {
        ranges,
        kind: "none".to_owned(),
        operator: 0,
        formula1: None,
        formula2: None,
        allow_blank: true,
        dropdown: true,
        show_input: true,
        show_error: true,
        error_style: "stop".to_owned(),
        prompt_title: None,
        prompt: None,
        error_title: None,
        error: None,
    };
    state.apply(props, Operation::Add)?;
    let existing = features::list(file, parent.sheet, Kind::DataValidation)?;
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
    let name = sheet_name(file, parent.sheet)?;
    let mut index = 0;
    let commit = publish(file, limits, |edit| {
        index = edit.add_sheet_validation(&name, &spec).map_err(native)?;
        Ok(())
    })?;
    Ok((
        commit,
        feature_path(parent.sheet, "dataValidation", Some(index)),
    ))
}

fn set_validation(
    file: &Ods,
    sheet: u32,
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
    let mut state = DvState::of(current);
    state.apply(props, Operation::Set)?;
    let index = feature.index as usize;
    let existing = features::list(file, sheet, Kind::DataValidation)?;
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
    let name = sheet_name(file, sheet)?;
    publish(file, limits, |edit| {
        edit.set_sheet_validation(&name, index, &spec)
            .map_err(native)
    })
}

// ---------------------------------------------------------------------------
// The auto-filter

fn filter_range(key: &'static str, value: &str) -> Result<Rect> {
    parse_rect(value).ok_or_else(|| invalid(key, "the range is not an A1 range"))
}

fn add_autofilter(
    file: &Ods,
    parent: &Parent,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    let mut range = parent.reference;
    for (key, name, arg) in admit(&features::AUTOFILTER_TABLE, Operation::Add, props)? {
        if key == AutofilterKey::Range {
            range = Some(filter_range(name, &arg.text().unwrap_or_default())?);
        }
    }
    let range = range.ok_or_else(|| missing("range", "an auto-filter needs a range"))?;
    let name = sheet_name(file, parent.sheet)?;
    let commit = publish(file, limits, |edit| {
        edit.set_sheet_auto_filter(&name, Some(range))
            .map_err(native)
    })?;
    Ok((commit, feature_path(parent.sheet, "autofilter", None)))
}

fn set_autofilter(
    file: &Ods,
    sheet: u32,
    feature: &Feature,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let Data::Autofilter(current) = &feature.data else {
        return Err(error(
            Code::InternalError,
            "an auto-filter node holds no filter",
        ));
    };
    let mut range = current.range;
    for (key, name, arg) in admit(&features::AUTOFILTER_TABLE, Operation::Set, props)? {
        if key == AutofilterKey::Range {
            range = filter_range(name, &arg.text().unwrap_or_default())?;
        }
    }
    let name = sheet_name(file, sheet)?;
    publish(file, limits, |edit| {
        edit.set_sheet_auto_filter(&name, Some(range))
            .map_err(native)
    })
}

// ---------------------------------------------------------------------------
// Page breaks

fn break_index(kind: Kind, key: &'static str, arg: &Arg) -> Result<u32> {
    let (maximum, letters) = if kind == Kind::RowBreak {
        (i64::from(MAX_ROW), false)
    } else {
        (i64::from(MAX_COLUMN), true)
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

fn breaks(file: &Ods, sheet: u32, kind: Kind) -> Result<Vec<u32>> {
    Ok(features::list(file, sheet, kind)?
        .into_iter()
        .filter_map(|feature| match feature.data {
            Data::Break(index) => Some(index),
            _ => None,
        })
        .collect())
}

fn publish_breaks(
    file: &Ods,
    sheet: u32,
    kind: Kind,
    breaks: &[u32],
    limits: &SemanticLimits,
) -> Result<Commit> {
    let name = sheet_name(file, sheet)?;
    publish(file, limits, |edit| {
        edit.set_sheet_breaks(&name, kind == Kind::RowBreak, breaks)
            .map_err(native)
    })
}

fn break_spec(
    kind: Kind,
    current: Option<u32>,
    props: &Props<'_>,
    operation: Operation,
) -> Result<u32> {
    let mut index = current;
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
                        "an ODS sheet stores manual page breaks only",
                    ));
                }
            },
            BreakKey::Min | BreakKey::Max => {},
        }
    }
    let key = if kind == Kind::RowBreak { "row" } else { "col" };
    index.ok_or_else(|| missing(key, "a page break needs its row or column"))
}

fn add_break(
    file: &Ods,
    parent: &Parent,
    kind: Kind,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    let item = break_spec(kind, None, props, Operation::Add)?;
    let mut held = breaks(file, parent.sheet, kind)?;
    if held.contains(&item) {
        return Err(error(
            Code::InvalidArgument,
            "the worksheet already has this page break",
        ));
    }
    let index = held.partition_point(|value| *value < item);
    held.insert(index, item);
    let commit = publish_breaks(file, parent.sheet, kind, &held, limits)?;
    Ok((
        commit,
        feature_path(parent.sheet, features::step_name(kind), Some(index)),
    ))
}

fn set_break(
    file: &Ods,
    sheet: u32,
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
    let item = break_spec(feature.kind, Some(current), props, Operation::Set)?;
    let mut held = breaks(file, sheet, feature.kind)?;
    let index = feature.index as usize;
    if index >= held.len() {
        return Err(features::missing("the page break is missing"));
    }
    held.remove(index);
    if held.contains(&item) {
        return Err(error(
            Code::InvalidArgument,
            "the worksheet already has this page break",
        ));
    }
    let at = held.partition_point(|value| *value < item);
    held.insert(at, item);
    publish_breaks(file, sheet, feature.kind, &held, limits)
}

/// The worksheet feature an element word names (`hyperlink`, `validation`,
/// `image`, `textbox`, a CF word, `pagebreak`), with the rule type a CF word
/// implies.
pub(super) fn feature_kind(name: &str, props: &Props<'_>) -> Option<(Kind, Option<&'static str>)> {
    if let Some(implied) = cf_element(name) {
        return Some((Kind::Cf, Some(implied)));
    }
    if name.eq_ignore_ascii_case("pagebreak") {
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
