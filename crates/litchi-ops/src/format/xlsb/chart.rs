//! XLSB worksheet charts, read-only (record 0322).
//!
//! An XLSB worksheet's charts are DrawingML chart parts its drawing's
//! graphic frames name, the same markup as XLSX's. `/Sheet1/chart[N]`,
//! `/Sheet1/chart[N]/series[K]` and `/Sheet1/chart[N]/axis[@role=ROLE]`
//! read them through the shared [`Summary`] projection with the XLSX read
//! vocabulary of record 0270: `chartType`, `title`, `legend`, `categories`,
//! `catTitle`, `seriesCount` and the anchor's cell geometry; a series'
//! `name`, `values`, `categories` and `color`; an axis' `role`, `title`,
//! `min` and `max`. Chart writes stay a gap (record 0322 follow-up).

use std::borrow::Cow;
use std::sync::Arc;

use litchi_core::{FileFormat, Position};
use litchi_xlsb as xlsb;
use xlsb::drawing::edit::Placement;

use super::node::{Item, Node, Sheet, SheetHandle};
use super::spec::Kind;
use super::{Xlsb, XlsbOwner};
use crate::error::{Code, Error, Result};
use crate::format::chart::{self as shared, LazySummary, Role, Summary};
use crate::node::spec::{GapSpec, PropSpec};
use crate::node::{PathStep, Value};
use crate::path::navigate::{Lookup, Select};
use crate::value::ValueType;

/// A worksheet chart's readable keys, by index.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ChartKey {
    ChartType,
    Title,
    Legend,
    Categories,
    CatTitle,
    SeriesCount,
    X,
    Y,
    Width,
    Height,
}

const CHART_KEYS: [ChartKey; 10] = [
    ChartKey::ChartType,
    ChartKey::Title,
    ChartKey::Legend,
    ChartKey::Categories,
    ChartKey::CatTitle,
    ChartKey::SeriesCount,
    ChartKey::X,
    ChartKey::Y,
    ChartKey::Width,
    ChartKey::Height,
];

/// A worksheet chart (`/Sheet1/chart[N]`), read from its chart part.
pub(crate) const CHART_PROPS: &[PropSpec] = &[
    PropSpec::new("chartType", ValueType::Text),
    PropSpec::new("title", ValueType::Text),
    PropSpec::new("legend", ValueType::Text),
    PropSpec::new("categories", ValueType::Text),
    PropSpec::new("catTitle", ValueType::Text),
    PropSpec::new("seriesCount", ValueType::Int),
    PropSpec::new("x", ValueType::Int),
    PropSpec::new("y", ValueType::Int),
    PropSpec::new("width", ValueType::Int),
    PropSpec::new("height", ValueType::Int),
];

const _: () = assert!(CHART_KEYS.len() == CHART_PROPS.len());

/// A chart series (`/Sheet1/chart[N]/series[K]`).
pub(crate) const SERIES_PROPS: &[PropSpec] = &[
    PropSpec::new("name", ValueType::Text),
    PropSpec::new("values", ValueType::Text),
    PropSpec::new("categories", ValueType::Text),
    PropSpec::new("color", ValueType::Text),
];

/// A chart axis (`/Sheet1/chart[N]/axis[@role=ROLE]`).
pub(crate) const AXIS_PROPS: &[PropSpec] = &[
    PropSpec::new("role", ValueType::Text),
    PropSpec::new("title", ValueType::Text),
    PropSpec::new("min", ValueType::Number),
    PropSpec::new("max", ValueType::Number),
];

/// A worksheet chart read from its part: the chart's position on the sheet,
/// its anchor's cell geometry and, for a series or axis, the child's
/// position in the summary.
#[derive(Clone)]
pub(crate) struct ChartItem {
    pub(super) chart: u32,
    part: Arc<LazySummary>,
    cells: Option<[u32; 4]>,
    pub(super) child: u32,
}

impl ChartItem {
    /// The chart part's projection, read on first use.
    pub(super) fn summary(&self) -> Result<&Arc<Summary>> {
        self.part.get()
    }

    fn with_child(&self, child: usize) -> Self {
        Self {
            chart: self.chart,
            part: Arc::clone(&self.part),
            cells: self.cells,
            child: u32::try_from(child).unwrap_or(u32::MAX),
        }
    }

    /// The node's last path step.
    pub(super) fn step(&self, kind: Kind) -> PathStep<'static> {
        match kind {
            Kind::Chart => PathStep::at("chart", Position::new(self.chart as usize)),
            Kind::Series => PathStep::at("series", Position::new(self.child as usize)),
            _ => PathStep::at("axis", Position::new(self.child as usize)),
        }
    }

    /// The chart's title, the text a query matches.
    pub(super) fn title(&self) -> Result<Option<Cow<'_, str>>> {
        Ok(self.summary()?.title.as_deref().map(Cow::Borrowed))
    }

    /// The value at `index` of the property table of `kind`.
    pub(super) fn value(&self, kind: Kind, index: usize) -> Option<Value<'_>> {
        let summary = self.summary().ok()?;
        match kind {
            Kind::Chart => chart_value(summary, self.cells, *CHART_KEYS.get(index)?),
            Kind::Series => {
                let series = summary.series.get(self.child as usize)?;
                match index {
                    0 => series
                        .name
                        .as_deref()
                        .map(|name| Value::Text(Cow::Borrowed(name))),
                    1 => Some(Value::Text(Cow::Owned(shared::join_values(&series.values)))),
                    2 => (!series.categories.is_empty())
                        .then(|| Value::Text(Cow::Owned(series.categories.join(",")))),
                    3 => series
                        .color
                        .as_deref()
                        .map(|color| Value::Text(Cow::Borrowed(color))),
                    _ => None,
                }
            },
            _ => {
                let axis = summary.axes.get(self.child as usize)?;
                match index {
                    0 => Some(Value::Text(Cow::Borrowed(axis.role.name()))),
                    1 => axis
                        .title
                        .as_deref()
                        .map(|title| Value::Text(Cow::Borrowed(title))),
                    2 => axis.min.and_then(Value::number),
                    3 => axis.max.and_then(Value::number),
                    _ => None,
                }
            },
        }
    }
}

/// The property table of a chart element kind.
pub(super) fn props(kind: Kind) -> &'static [PropSpec] {
    match kind {
        Kind::Chart => CHART_PROPS,
        Kind::Series => SERIES_PROPS,
        _ => AXIS_PROPS,
    }
}

fn chart_value(summary: &Summary, cells: Option<[u32; 4]>, key: ChartKey) -> Option<Value<'_>> {
    match key {
        ChartKey::ChartType => summary
            .chart_type()
            .map(|name| Value::Text(Cow::Owned(name))),
        ChartKey::Title => summary
            .title
            .as_deref()
            .map(|title| Value::Text(Cow::Borrowed(title))),
        ChartKey::Legend => Some(Value::Text(Cow::Borrowed(shared::legend_name(
            summary.legend,
        )))),
        ChartKey::Categories => summary
            .categories()
            .map(|categories| Value::Text(Cow::Owned(categories.join(",")))),
        ChartKey::CatTitle => summary
            .axis(Role::Category)
            .and_then(|axis| axis.title.as_deref())
            .map(|title| Value::Text(Cow::Borrowed(title))),
        ChartKey::SeriesCount => Some(Value::Int(
            i64::try_from(summary.series.len()).unwrap_or(i64::MAX),
        )),
        ChartKey::X | ChartKey::Y | ChartKey::Width | ChartKey::Height => {
            let cells = cells?;
            let slot = match key {
                ChartKey::X => 0,
                ChartKey::Y => 1,
                ChartKey::Width => 2,
                _ => 3,
            };
            Some(Value::Int(i64::from(cells[slot])))
        },
    }
}

/// OfficeCLI's cell geometry of a two-cell anchor: `[x, y, width, height]`.
fn cells(placement: Option<Placement>) -> Option<[u32; 4]> {
    match placement? {
        Placement::TwoCell { from, to } => Some([
            from.col,
            from.row,
            to.col.saturating_sub(from.col),
            to.row.saturating_sub(from.row),
        ]),
        _ => None,
    }
}

fn lost_owner() -> Error {
    Error::new(
        Code::InternalError,
        "an XLSB worksheet handle lost its owner state",
    )
    .with_format(FileFormat::Xlsb)
}

impl SheetHandle {
    /// The worksheet's charts, in drawing order.
    fn charts(&self) -> Result<Vec<xlsb::drawing::edit::ChartPart>> {
        if !self.is_grid() {
            return Ok(Vec::new());
        }
        Ok(match self {
            Self::Eager { file, .. } => match &file.0.owner {
                XlsbOwner::Eager(package) => {
                    super::projection(package)?.drawing_charts(self.position().get())?
                },
                XlsbOwner::Source(_) => return Err(lost_owner()),
            },
            Self::Source { sheet, .. } => sheet.drawing_charts()?,
        })
    }
}

/// The charts of a worksheet, in drawing order.
pub(super) fn list(sheet: &Sheet) -> Result<Vec<ChartItem>> {
    Ok(sheet
        .handle
        .charts()?
        .into_iter()
        .enumerate()
        .map(|(index, part)| ChartItem {
            chart: u32::try_from(index).unwrap_or(u32::MAX),
            cells: cells(part.placement),
            part: Arc::new(LazySummary::new(part.xml)),
            child: 0,
        })
        .collect())
}

/// Resolve a chart step under a worksheet.
pub(super) fn lookup<'f>(
    file: &'f Xlsb,
    sheet: &Sheet,
    select: Select<'_>,
) -> Result<Lookup<Node<'f>, GapSpec>> {
    let index = match select {
        Select::Ordinal(position) => Some(position.get()),
        Select::None => Some(0),
        Select::Last => None,
        _ => return Ok(Lookup::NotApplicable),
    };
    let mut charts = list(sheet)?;
    let index = index.or_else(|| charts.len().checked_sub(1));
    Ok(match index.filter(|index| *index < charts.len()) {
        Some(index) => Lookup::Found(Node::new_chart(
            file,
            sheet.position,
            Kind::Chart,
            charts.swap_remove(index),
        )),
        None => Lookup::Missing,
    })
}

/// A worksheet chart's series (by position) or axis (by position or
/// `@role`).
pub(super) fn child<'f>(
    file: &'f Xlsb,
    parent: &Node<'f>,
    chart: &ChartItem,
    kind: Kind,
    select: Select<'_>,
) -> Result<Lookup<Node<'f>, GapSpec>> {
    let summary = chart.summary()?;
    let count = if kind == Kind::Series {
        summary.series.len()
    } else {
        summary.axes.len()
    };
    let index = match select {
        Select::Ordinal(ordinal) => Some(ordinal.get()),
        Select::None => Some(0),
        Select::Last => count.checked_sub(1),
        Select::Identity(identities) if kind == Kind::Axis => {
            let mut found = None;
            for (identity, value) in identities.iter() {
                if !identity.key.eq_ignore_ascii_case("role") {
                    return Ok(Lookup::NotApplicable);
                }
                found = Role::parse(value)
                    .and_then(|role| summary.axes.iter().position(|axis| axis.role == role));
            }
            found
        },
        _ => return Ok(Lookup::NotApplicable),
    };
    Ok(match index.filter(|index| *index < count) {
        Some(index) => Lookup::Found(Node::new_chart(
            file,
            parent.sheet_position(),
            kind,
            chart.with_child(index),
        )),
        None => Lookup::Missing,
    })
}

/// A chart's series and axes, as wanted.
pub(super) fn parts(chart: &ChartItem, series: bool, axes: bool) -> Result<Vec<(Kind, ChartItem)>> {
    let summary = chart.summary()?;
    let mut items = Vec::new();
    if series {
        items
            .extend((0..summary.series.len()).map(|index| (Kind::Series, chart.with_child(index))));
    }
    if axes {
        items.extend((0..summary.axes.len()).map(|index| (Kind::Axis, chart.with_child(index))));
    }
    Ok(items)
}

impl<'f> Node<'f> {
    /// A worksheet chart, or one of its series or axes.
    pub(super) fn new_chart(file: &'f Xlsb, sheet: usize, kind: Kind, item: ChartItem) -> Self {
        let item = match kind {
            Kind::Chart => Item::Chart(Box::new(item)),
            Kind::Series => Item::Series(Box::new(item)),
            _ => Item::Axis(Box::new(item)),
        };
        Self::in_sheet_item(file, sheet, item)
    }
}
