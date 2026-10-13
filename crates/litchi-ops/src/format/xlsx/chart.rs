//! Typed worksheet chart CRUD (record 0270).
//!
//! - `add(/Sheet1, chart, {chartType, title, data | dataRange, categories,
//!   legend, catTitle, axisTitle, name, x, y, width, height})` builds an
//!   authored chart spec ([`crate::format::chart`]) and places it through
//!   the leaf chart family (`litchi_xlsx::chart::edit`). Inline `data` is
//!   written as literals, so no workbook part is needed; `dataRange`
//!   (`Sheet1!A1:E13`: series names in the first row, categories in the
//!   first column) reads the host cells and refers to them, with cached
//!   values. `x`/`y`/`width`/`height` count cells, as OfficeCLI's do.
//! - `get` reads `/Sheet1/chart[N]`, `/Sheet1/chart[N]/series[K]` and
//!   `/Sheet1/chart[N]/axis[@role=ROLE]` from the chart part.
//! - `set` edits the chart's title, legend, categories and category-axis
//!   title, a series' name and values, and an axis title through one
//!   source-preserving chart edit; `remove` deletes the chart's anchor and
//!   part.

use std::borrow::Cow;
use std::sync::Arc;

use litchi_core::selector::{Path, Predicate, Step};
use litchi_core::{FileFormat, Position};
use litchi_drawingml::chart::authoring::{Binding, Series};
use litchi_xlsx as xlsx;

use super::node::{ChartItem, Item, Node};
use super::write::Commit;
use super::{Nav, Xlsx};
use crate::error::{Code, Error, Result};
use crate::format::chart::{self as shared, AddProps, ChartEdit, Role, Summary};
use crate::new_node::NewNode;
use crate::node::Value;
use crate::node::spec::PropSpec;
use crate::placement::Placement;
use crate::props::Props;
use crate::registry::{Access, Alias, Constraint, Entry, Operation, Table};
use crate::target::TargetArg;
use crate::value::{Scalar, ValueType};

fn error(code: Code, message: &'static str) -> Error {
    Error::new(code, message).with_format(FileFormat::Xlsx)
}

const fn write_only(name: &'static str, ty: ValueType) -> PropSpec {
    let mut spec = PropSpec::new(name, ty);
    spec.access.get = false;
    spec
}

/// A worksheet chart's readable keys, by index.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum ChartKey {
    ChartType,
    Title,
    Legend,
    Categories,
    CatTitle,
    SeriesCount,
}

impl ChartKey {
    const ALL: [Self; 6] = [
        Self::ChartType,
        Self::Title,
        Self::Legend,
        Self::Categories,
        Self::CatTitle,
        Self::SeriesCount,
    ];

    pub(crate) fn at(index: usize) -> Option<Self> {
        Self::ALL.get(index).copied()
    }
}

/// A worksheet chart (`/Sheet1/chart[N]`), read from its chart part.
pub(crate) const CHART_PROPS: &[PropSpec] = &[
    PropSpec::new("chartType", ValueType::Text),
    PropSpec::new("title", ValueType::Text),
    PropSpec::new("legend", ValueType::Text),
    PropSpec::new("categories", ValueType::Text),
    PropSpec::new("catTitle", ValueType::Text),
    PropSpec::new("seriesCount", ValueType::Int),
];

const _: () = assert!(ChartKey::ALL.len() == CHART_PROPS.len());

/// A chart series' keys.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum SeriesKey {
    Name,
    Values,
    Categories,
}

/// A chart series (`/Sheet1/chart[N]/series[K]`).
pub(crate) const SERIES_PROPS: &[PropSpec] = &[
    PropSpec::new("name", ValueType::Text),
    PropSpec::new("values", ValueType::Text),
    PropSpec::new("categories", ValueType::Text),
];

/// A chart axis (`/Sheet1/chart[N]/axis[@role=ROLE]`).
pub(crate) const AXIS_PROPS: &[PropSpec] = &[
    PropSpec::new("role", ValueType::Text),
    PropSpec::new("title", ValueType::Text),
];

/// Chart assignments.
#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) enum ChartSetter {
    ChartType,
    Title,
    Legend,
    Categories,
    CatTitle,
    Data,
    DataRange,
    AxisTitle,
    Name,
    X,
    Y,
    Width,
    Height,
    /// A formatting key ([`shared::FORMAT_PROPS`]).
    Format(shared::FormatKey),
    /// `series{N}`, by its zero-based number.
    Numbered(u8),
}

const CELL_SPAN: Constraint = Constraint::Int {
    min: 1,
    max: 16_384,
};
const CELL_ORIGIN: Constraint = Constraint::Int {
    min: 0,
    max: 1_048_575,
};

const fn entry(
    spec: PropSpec,
    access: Access<ChartSetter>,
    constraint: Constraint,
    clear: bool,
) -> Entry<ChartSetter> {
    Entry {
        spec,
        access,
        constraint,
        clear,
    }
}

const CHART_BASE_COUNT: usize = 14;
const CHART_FORMAT_START: usize = CHART_BASE_COUNT;
const CHART_NUMBERED_START: usize = CHART_FORMAT_START + shared::FORMAT_PROPS.len();
const CHART_WRITE_COUNT: usize = CHART_NUMBERED_START + shared::NUMBERED_SERIES_PROPS.len();
const _: () = assert!(
    CHART_WRITE_COUNT <= 64,
    "chart adds track assignments in a u64"
);

const fn chart_entries() -> [Entry<ChartSetter>; CHART_WRITE_COUNT] {
    let base: [Entry<ChartSetter>; CHART_BASE_COUNT] = [
        entry(
            CHART_PROPS[0],
            Access::Add(ChartSetter::ChartType),
            Constraint::None,
            false,
        ),
        entry(
            CHART_PROPS[1],
            Access::SetAndAdd(ChartSetter::Title),
            Constraint::None,
            true,
        ),
        entry(
            CHART_PROPS[2],
            Access::SetAndAdd(ChartSetter::Legend),
            Constraint::None,
            false,
        ),
        entry(
            CHART_PROPS[3],
            Access::SetAndAdd(ChartSetter::Categories),
            Constraint::NotEmpty,
            false,
        ),
        entry(
            CHART_PROPS[4],
            Access::SetAndAdd(ChartSetter::CatTitle),
            Constraint::None,
            true,
        ),
        Entry::read(CHART_PROPS[5]),
        entry(
            write_only("data", ValueType::Text),
            Access::Add(ChartSetter::Data),
            Constraint::NotEmpty,
            false,
        ),
        entry(
            write_only("dataRange", ValueType::Text),
            Access::Add(ChartSetter::DataRange),
            Constraint::NotEmpty,
            false,
        ),
        entry(
            write_only("axisTitle", ValueType::Text),
            Access::Add(ChartSetter::AxisTitle),
            Constraint::None,
            false,
        ),
        entry(
            write_only("name", ValueType::Text),
            Access::Add(ChartSetter::Name),
            Constraint::None,
            false,
        ),
        entry(
            write_only("x", ValueType::Int),
            Access::Add(ChartSetter::X),
            CELL_ORIGIN,
            false,
        ),
        entry(
            write_only("y", ValueType::Int),
            Access::Add(ChartSetter::Y),
            CELL_ORIGIN,
            false,
        ),
        entry(
            write_only("width", ValueType::Int),
            Access::Add(ChartSetter::Width),
            CELL_SPAN,
            false,
        ),
        entry(
            write_only("height", ValueType::Int),
            Access::Add(ChartSetter::Height),
            Constraint::Int {
                min: 1,
                max: 1_048_576,
            },
            false,
        ),
    ];
    let mut entries = [base[0]; CHART_WRITE_COUNT];
    let mut index = 0;
    while index < CHART_BASE_COUNT {
        entries[index] = base[index];
        index += 1;
    }
    let mut format = 0;
    while format < shared::FORMAT_PROPS.len() {
        let (key, spec) = shared::FORMAT_PROPS[format];
        entries[CHART_FORMAT_START + format] = entry(
            spec,
            Access::Add(ChartSetter::Format(key)),
            Constraint::None,
            false,
        );
        format += 1;
    }
    let mut number = 0;
    while number < shared::NUMBERED_SERIES_PROPS.len() {
        entries[CHART_NUMBERED_START + number] = entry(
            shared::NUMBERED_SERIES_PROPS[number],
            Access::Add(ChartSetter::Numbered(number as u8)),
            Constraint::NotEmpty,
            false,
        );
        number += 1;
    }
    entries
}

pub(crate) const CHART_WRITE_TABLE: Table<ChartSetter, CHART_WRITE_COUNT, 3> = Table::new(
    chart_entries(),
    [
        Alias {
            name: "type",
            entry: 0,
        },
        Alias {
            name: "htitle",
            entry: 4,
        },
        Alias {
            name: "vtitle",
            entry: 8,
        },
    ],
);

pub(crate) const CHART_CAPABILITY_PROPS: &[PropSpec] = &CHART_WRITE_TABLE.readable();

#[derive(Clone, Copy)]
pub(crate) enum SeriesSetter {
    Name,
    Values,
}
pub(crate) const SERIES_WRITE_TABLE: Table<SeriesSetter, 3, 0> = Table::new(
    [
        Entry {
            spec: SERIES_PROPS[0],
            access: Access::Set(SeriesSetter::Name),
            constraint: Constraint::None,
            clear: false,
        },
        Entry {
            spec: SERIES_PROPS[1],
            access: Access::Set(SeriesSetter::Values),
            constraint: Constraint::NotEmpty,
            clear: false,
        },
        Entry::read(SERIES_PROPS[2]),
    ],
    [],
);
pub(crate) const SERIES_CAPABILITY_PROPS: &[PropSpec] = &SERIES_WRITE_TABLE.readable();

#[derive(Clone, Copy)]
pub(crate) enum AxisSetter {
    Title,
}
pub(crate) const AXIS_WRITE_TABLE: Table<AxisSetter, 2, 0> = Table::new(
    [
        Entry::read(AXIS_PROPS[0]),
        Entry {
            spec: AXIS_PROPS[1],
            access: Access::Set(AxisSetter::Title),
            constraint: Constraint::None,
            clear: true,
        },
    ],
    [],
);
pub(crate) const AXIS_CAPABILITY_PROPS: &[PropSpec] = &AXIS_WRITE_TABLE.readable();

fn text(value: Option<Scalar<'_>>) -> Option<String> {
    match value {
        Some(Scalar::Text(value) | Scalar::Token(value)) => Some(value.to_owned()),
        _ => None,
    }
}

fn int(value: Option<Scalar<'_>>) -> u32 {
    match value {
        Some(Scalar::Int(value)) => u32::try_from(value).unwrap_or(0),
        _ => 0,
    }
}

/// The value of a chart node's key.
pub(crate) fn chart_value(summary: &Summary, key: ChartKey) -> Option<Value<'static>> {
    match key {
        ChartKey::ChartType => summary
            .chart_type()
            .map(|name| Value::Text(Cow::Owned(name))),
        ChartKey::Title => summary
            .title
            .clone()
            .map(|title| Value::Text(Cow::Owned(title))),
        ChartKey::Legend => Some(Value::Text(Cow::Borrowed(shared::legend_name(
            summary.legend,
        )))),
        ChartKey::Categories => summary
            .categories()
            .map(|categories| Value::Text(Cow::Owned(categories.join(",")))),
        ChartKey::CatTitle => summary
            .axis(Role::Category)
            .and_then(|axis| axis.title.clone())
            .map(|title| Value::Text(Cow::Owned(title))),
        ChartKey::SeriesCount => Some(Value::Int(
            i64::try_from(summary.series.len()).unwrap_or(i64::MAX),
        )),
    }
}

/// The value of a series node's key (`name`, `values`, `categories`).
pub(crate) fn series_value(item: &ChartItem, index: usize) -> Option<Value<'_>> {
    let series = item.summary.series.get(item.child as usize)?;
    let key = [SeriesKey::Name, SeriesKey::Values, SeriesKey::Categories].get(index)?;
    match key {
        SeriesKey::Name => series
            .name
            .as_deref()
            .map(|name| Value::Text(Cow::Borrowed(name))),
        SeriesKey::Values => Some(Value::Text(Cow::Owned(shared::join_values(&series.values)))),
        SeriesKey::Categories => (!series.categories.is_empty())
            .then(|| Value::Text(Cow::Owned(series.categories.join(",")))),
    }
}

/// The value of an axis node's key (`role`, `title`).
pub(crate) fn axis_value(item: &ChartItem, index: usize) -> Option<Value<'_>> {
    let axis = item.summary.axes.get(item.child as usize)?;
    match index {
        0 => Some(Value::Text(Cow::Borrowed(axis.role.name()))),
        1 => axis
            .title
            .as_deref()
            .map(|title| Value::Text(Cow::Borrowed(title))),
        _ => None,
    }
}

impl Xlsx {
    /// The number of ordinary charts on the worksheet at `sheet`.
    pub(super) fn sheet_chart_count(&self, sheet: usize) -> Result<usize> {
        self.0.owner.finish(|| {
            Ok(match &self.0.owner {
                super::XlsxOwner::Source(workbook) => match workbook.sheet(sheet)? {
                    Some(worksheet) if worksheet.kind() == xlsx::WorksheetKind::Worksheet => {
                        worksheet.chart_count()?
                    },
                    _ => 0,
                },
                super::XlsxOwner::Eager(package) => match package.workbook_ref().sheet(sheet)? {
                    Some(worksheet) if worksheet.kind() == xlsx::WorksheetKind::Worksheet => {
                        package.chart_count(sheet)?
                    },
                    _ => 0,
                },
            })
        })
    }

    /// The chart at `index` on the worksheet at `sheet`, read from its part.
    pub(super) fn sheet_chart(&self, sheet: usize, index: usize) -> Result<Option<Arc<Summary>>> {
        let xml = self.0.owner.finish(|| {
            Ok(match &self.0.owner {
                super::XlsxOwner::Source(workbook) => match workbook.sheet(sheet)? {
                    Some(worksheet) if worksheet.kind() == xlsx::WorksheetKind::Worksheet => {
                        worksheet.chart_xml(index)?
                    },
                    _ => None,
                },
                super::XlsxOwner::Eager(package) => match package.workbook_ref().sheet(sheet)? {
                    Some(worksheet) if worksheet.kind() == xlsx::WorksheetKind::Worksheet => {
                        package.chart_xml(sheet, index)?.map(<[u8]>::to_vec)
                    },
                    _ => None,
                },
            })
        })?;
        xml.map(|xml| Summary::read(&xml).map(Arc::new)).transpose()
    }

    /// Add a chart to a worksheet.
    pub(super) fn add_chart(
        &self,
        parent: &Node<'_>,
        node: NewNode<'_>,
    ) -> Result<(Commit, Path<'static>)> {
        let Item::Sheet(sheet) = parent.item() else {
            return Err(error(
                Code::UnsupportedType,
                "charts require a worksheet parent",
            ));
        };
        let sheet = sheet.position;
        if !matches!(node.placement(), None | Some(Placement::End)) {
            return Err(error(
                Code::UnsupportedCapability,
                "chart insertion currently supports end placement",
            ));
        }
        let mut props = AddProps::default();
        let mut range = None;
        let mut name = None;
        let mut origin = [0_u32; 2];
        let mut span = [8_u32, 15];
        let mut assigned = 0_u64;
        for (key, value) in node.props().iter() {
            let (index, entry) = CHART_WRITE_TABLE
                .resolve(key.as_str())
                .map_err(|e| e.with_property(key.as_str().to_owned()))?;
            let bit = 1_u64 << index;
            if assigned & bit != 0 {
                return Err(error(
                    Code::InvalidArgument,
                    "a chart property was assigned more than once",
                )
                .with_property(entry.spec.name));
            }
            assigned |= bit;
            entry.prepare(Operation::Add, value, |setter, scalar| {
                match setter {
                    ChartSetter::ChartType => {
                        props.variant =
                            Some(shared::parse_variant(&text(scalar).unwrap_or_default())?);
                    },
                    ChartSetter::Title => props.title = text(scalar),
                    ChartSetter::Legend => {
                        props.legend =
                            Some(shared::parse_legend(&text(scalar).unwrap_or_default())?);
                    },
                    ChartSetter::Categories => {
                        props.categories = Some(shared::parse_list(
                            &text(scalar).unwrap_or_default(),
                            "categories",
                        )?);
                    },
                    ChartSetter::CatTitle => props.category_axis_title = text(scalar),
                    ChartSetter::AxisTitle => props.value_axis_title = text(scalar),
                    ChartSetter::Data => {
                        props.data = Some(shared::parse_data(&text(scalar).unwrap_or_default())?);
                    },
                    ChartSetter::DataRange => range = text(scalar),
                    ChartSetter::Name => {
                        let value = text(scalar).unwrap_or_default();
                        if value.contains('\0') {
                            return Err(error(Code::InvalidValue, "the name contains NUL")
                                .with_property("name"));
                        }
                        name = Some(value);
                    },
                    ChartSetter::X => origin[0] = int(scalar),
                    ChartSetter::Y => origin[1] = int(scalar),
                    ChartSetter::Width => span[0] = int(scalar),
                    ChartSetter::Height => span[1] = int(scalar),
                    ChartSetter::Format(key) => props.format(key, scalar)?,
                    ChartSetter::Numbered(number) => props.numbered_series(
                        usize::from(number) + 1,
                        &text(scalar).unwrap_or_default(),
                    )?,
                }
                Ok(())
            })?;
        }
        let binding = match range {
            Some(range) => {
                if props.data.is_some() {
                    return Err(error(
                        Code::InvalidArgument,
                        "a chart takes data or dataRange, not both",
                    )
                    .with_property("dataRange"));
                }
                Some(self.read_range(sheet, &range, &mut props)?)
            },
            None => None,
        };
        let mut spec = props.into_spec()?;
        if let Some(binding) = binding {
            spec = spec.with_binding(binding);
            spec.validate().map_err(|error| {
                Error::new(Code::InvalidValue, error.to_string()).with_property("dataRange")
            })?;
        }
        let xml = spec.to_xml(None).map_err(|error| {
            Error::new(Code::InvalidValue, error.to_string()).with_property("data")
        })?;
        let to_col = origin[0].saturating_add(span[0]);
        let to_row = origin[1].saturating_add(span[1]);
        if to_col >= 16_384 || to_row >= 1_048_576 {
            return Err(
                error(Code::OutOfRange, "the chart extends past the worksheet")
                    .with_property("width"),
            );
        }
        let anchor = xlsx::chart::Anchor::new(origin[0], origin[1], to_col, to_row);
        let source = self.editor()?;
        let name = match name.or_else(|| spec.title().map(str::to_owned)) {
            Some(name) => name,
            // Counting reads the sheet's drawing: only for an unnamed chart.
            None => format!("Chart {}", source.chart_count(sheet)? + 1),
        };
        let mut edit = source.edit_charts()?;
        let index = edit.add(sheet, &anchor, &name, &xml)?;
        let target = edit.commit()?.into_package();
        let after = self.adopt_workbook(target.clone());
        Ok((
            Commit::from_metadata(after, source, target),
            Path::from_steps([
                Step::new("sheet").with(Predicate::Index(Position::new(sheet))),
                Step::new("chart").with(Predicate::Index(Position::new(index))),
            ]),
        ))
    }

    /// Read `Sheet!A1:E13`: series names in the first row, categories in the
    /// first column, values below; the chart refers to these cells.
    fn read_range(&self, sheet: usize, range: &str, props: &mut AddProps) -> Result<Binding> {
        let invalid = |message: &'static str| {
            Error::new(Code::InvalidValue, message)
                .with_property("dataRange")
                .with_format(FileFormat::Xlsx)
        };
        let (sheet_name, cells) = match range.rsplit_once('!') {
            Some((name, cells)) => (name.trim_matches('\'').replace("''", "'"), cells),
            None => (super::node::sheet_name(self, sheet).to_owned(), range),
        };
        let rect = cells
            .replace('$', "")
            .parse::<xlsx::Rect>()
            .map_err(|_| invalid("dataRange is not an A1 range"))?;
        let package = self.editor()?;
        let workbook = package.workbook();
        let worksheet = workbook
            .sheet(sheet_name.as_str())?
            .ok_or_else(|| invalid("dataRange names no worksheet"))?;
        let (first_row, first_column) = (rect.start().row().get(), rect.start().column().get());
        // `end` is exclusive.
        let (last_row, last_column) = (
            rect.end().0.saturating_sub(1),
            rect.end().1.saturating_sub(1),
        );
        if last_row <= first_row || last_column <= first_column {
            return Err(invalid(
                "dataRange needs a header row, a category column and at least one value",
            ));
        }
        let cell = |row: u32, column: u32| -> Result<(Option<String>, Option<f64>)> {
            let address = xlsx::Address::at(row, column)
                .map_err(|_| invalid("dataRange is outside the worksheet"))?;
            Ok(match worksheet.cell(address)?.stored() {
                Some(xlsx::Cell::Value(value)) => cell_value(value),
                Some(xlsx::Cell::Formula(formula)) => formula
                    .cached()
                    .map_or((None, None), |cache| cell_value(cache.value())),
                _ => (None, None),
            })
        };
        let mut categories = Vec::new();
        for row in first_row + 1..=last_row {
            let (label, number) = cell(row, first_column)?;
            categories.push(
                label
                    .or_else(|| number.map(|value| value.to_string()))
                    .unwrap_or_default(),
            );
        }
        let mut series = Vec::new();
        for column in first_column + 1..=last_column {
            let (header, number) = cell(first_row, column)?;
            let name = header
                .or_else(|| number.map(|value| value.to_string()))
                .unwrap_or_else(|| format!("Series {}", column - first_column));
            let mut values = Vec::new();
            for row in first_row + 1..=last_row {
                values.push(cell(row, column)?.1.unwrap_or(0.0));
            }
            series.push(Series::new(name, values));
        }
        if props.categories.is_none() {
            props.categories = Some(categories);
        }
        props.data = Some(series);
        Ok(Binding::Sheet {
            sheet: worksheet.name().to_owned(),
            row: first_row as usize,
            column: first_column as usize,
        })
    }

    /// A `set` on a chart, series or axis path; `None` for other targets.
    pub(super) fn try_set_chart(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
    ) -> Result<Option<Commit>> {
        let Some(path) = target.path() else {
            return Ok(None);
        };
        if !path.path().steps().last().is_some_and(|step| {
            ["chart", "series", "axis"]
                .iter()
                .any(|name| step.name().eq_ignore_ascii_case(name))
        }) {
            return Ok(None);
        }
        let node = crate::path::resolve::get(&Nav::new(self), path)?;
        let (Item::Chart(item) | Item::Series(item) | Item::Axis(item)) = node.item() else {
            return Ok(None);
        };
        let mut edit = ChartEdit::default();
        for (key, value) in props.iter() {
            match node.item() {
                Item::Chart(_) => {
                    let (_, entry) = CHART_WRITE_TABLE
                        .resolve(key.as_str())
                        .map_err(|e| e.with_property(key.as_str().to_owned()))?;
                    entry.prepare(Operation::Set, value, |setter, scalar| {
                        match setter {
                            ChartSetter::Title => edit.title = Some(text(scalar)),
                            ChartSetter::Legend => {
                                edit.legend =
                                    Some(shared::parse_legend(&text(scalar).unwrap_or_default())?);
                            },
                            ChartSetter::Categories => {
                                edit.categories = Some(shared::parse_list(
                                    &text(scalar).unwrap_or_default(),
                                    "categories",
                                )?);
                            },
                            ChartSetter::CatTitle => {
                                let axis = item.summary.axis(Role::Category).ok_or_else(|| {
                                    error(
                                        Code::UnsupportedCapability,
                                        "the chart has no category axis",
                                    )
                                    .with_property("catTitle")
                                })?;
                                edit.axis_titles.push((axis.id, text(scalar)));
                            },
                            _ => unreachable!("chart set admission"),
                        }
                        Ok(())
                    })?;
                },
                Item::Series(_) => {
                    let index = item
                        .summary
                        .series
                        .get(item.child as usize)
                        .map(|series| series.index)
                        .ok_or_else(|| error(Code::PathNotFound, "the series is missing"))?;
                    let (_, entry) = SERIES_WRITE_TABLE
                        .resolve(key.as_str())
                        .map_err(|e| e.with_property(key.as_str().to_owned()))?;
                    entry.prepare(Operation::Set, value, |setter, scalar| {
                        match setter {
                            SeriesSetter::Name => edit.series_names.push((index, text(scalar))),
                            SeriesSetter::Values => edit.series_values.push((
                                index,
                                shared::parse_values(&text(scalar).unwrap_or_default(), "values")?,
                            )),
                        }
                        Ok(())
                    })?;
                },
                _ => {
                    let id = item
                        .summary
                        .axes
                        .get(item.child as usize)
                        .map(|axis| axis.id)
                        .ok_or_else(|| error(Code::PathNotFound, "the axis is missing"))?;
                    let (_, entry) = AXIS_WRITE_TABLE
                        .resolve(key.as_str())
                        .map_err(|e| e.with_property(key.as_str().to_owned()))?;
                    entry.prepare(Operation::Set, value, |setter, scalar| {
                        match setter {
                            AxisSetter::Title => edit.axis_titles.push((id, text(scalar))),
                        }
                        Ok(())
                    })?;
                },
            }
        }
        let sheet = node.sheet_position();
        let source = self.editor()?;
        let mut native = source.edit_charts()?;
        let xml = native.chart(sheet, item.chart as usize)?;
        if let Some(edited) = edit.apply(xml)? {
            native.set(sheet, item.chart as usize, edited.xml)?;
        }
        let commit = native.commit()?;
        let changed = commit.diagnostics().changed();
        let target = commit.into_package();
        let after = if changed {
            self.adopt_workbook(target.clone())
        } else {
            self.clone()
        };
        Ok(Some(Commit::from_metadata(after, source, target)))
    }

    /// Remove a chart selected by its path; `None` for other targets.
    pub(super) fn try_remove_chart(&self, target: &TargetArg<'_>) -> Result<Option<Commit>> {
        let Some(path) = target.path() else {
            return Ok(None);
        };
        if !path
            .path()
            .steps()
            .last()
            .is_some_and(|step| step.name().eq_ignore_ascii_case("chart"))
        {
            return Ok(None);
        }
        let node = crate::path::resolve::get(&Nav::new(self), path)?;
        let Item::Chart(item) = node.item() else {
            return Ok(None);
        };
        let sheet = node.sheet_position();
        let source = self.editor()?;
        let mut native = source.edit_charts()?;
        native.remove(sheet, item.chart as usize)?;
        let target = native.commit()?.into_package();
        let after = self.adopt_workbook(target.clone());
        Ok(Some(Commit::from_metadata(after, source, target)))
    }
}

fn cell_value(value: &xlsx::Value) -> (Option<String>, Option<f64>) {
    match value {
        xlsx::Value::Number(number) => (None, number.as_f64()),
        xlsx::Value::Text(text) => {
            let text = text.as_str();
            (Some(text.to_owned()), text.trim().parse::<f64>().ok())
        },
        xlsx::Value::Bool(value) => (Some(if *value { "TRUE" } else { "FALSE" }.to_owned()), None),
        _ => (None, None),
    }
}
