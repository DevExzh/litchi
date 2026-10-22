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
    Color,
}

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
    /// `series{N}.name`, `.values` or `.color`, by the zero-based number
    /// (record 0271).
    Dotted(u8, shared::SeriesField),
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
const CHART_DOTTED_START: usize = CHART_NUMBERED_START + shared::NUMBERED_SERIES_PROPS.len();
const CHART_WRITE_COUNT: usize = CHART_DOTTED_START + shared::DOTTED_SERIES_PROPS.len();
const _: () = assert!(
    CHART_WRITE_COUNT <= shared::MAX_TABLE,
    "chart adds track assignments in an Assigned set"
);

const fn chart_entries() -> [Entry<ChartSetter>; CHART_WRITE_COUNT] {
    let base: [Entry<ChartSetter>; CHART_BASE_COUNT] = [
        entry(
            CHART_PROPS[0],
            Access::SetAndAdd(ChartSetter::ChartType),
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
            Access::SetAndAdd(ChartSetter::Data),
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
            Access::SetAndAdd(ChartSetter::AxisTitle),
            Constraint::None,
            true,
        ),
        entry(
            write_only("name", ValueType::Text),
            Access::Add(ChartSetter::Name),
            Constraint::None,
            false,
        ),
        entry(
            PropSpec::new("x", ValueType::Int),
            Access::SetAndAdd(ChartSetter::X),
            CELL_ORIGIN,
            false,
        ),
        entry(
            PropSpec::new("y", ValueType::Int),
            Access::SetAndAdd(ChartSetter::Y),
            CELL_ORIGIN,
            false,
        ),
        entry(
            PropSpec::new("width", ValueType::Int),
            Access::SetAndAdd(ChartSetter::Width),
            CELL_SPAN,
            false,
        ),
        entry(
            PropSpec::new("height", ValueType::Int),
            Access::SetAndAdd(ChartSetter::Height),
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
            Access::SetAndAdd(ChartSetter::Format(key)),
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
    let mut dotted = 0;
    while dotted < shared::DOTTED_SERIES_PROPS.len() {
        let field = shared::SERIES_FIELDS[dotted % 3];
        let setter = ChartSetter::Dotted((dotted / 3) as u8, field);
        entries[CHART_DOTTED_START + dotted] = entry(
            shared::DOTTED_SERIES_PROPS[dotted],
            // A series' name, values and color are also settable on the
            // chart (OfficeCLI's chart-level `set`).
            Access::SetAndAdd(setter),
            Constraint::NotEmpty,
            false,
        );
        dotted += 1;
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
    Color,
}
pub(crate) const SERIES_WRITE_TABLE: Table<SeriesSetter, 4, 0> = Table::new(
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
        Entry {
            spec: SERIES_PROPS[3],
            access: Access::Set(SeriesSetter::Color),
            constraint: Constraint::NotEmpty,
            clear: true,
        },
    ],
    [],
);
pub(crate) const SERIES_CAPABILITY_PROPS: &[PropSpec] = &SERIES_WRITE_TABLE.readable();

#[derive(Clone, Copy)]
pub(crate) enum AxisSetter {
    Title,
    Min,
    Max,
}
pub(crate) const AXIS_WRITE_TABLE: Table<AxisSetter, 4, 0> = Table::new(
    [
        Entry::read(AXIS_PROPS[0]),
        Entry {
            spec: AXIS_PROPS[1],
            access: Access::Set(AxisSetter::Title),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: AXIS_PROPS[2],
            access: Access::Set(AxisSetter::Min),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: AXIS_PROPS[3],
            access: Access::Set(AxisSetter::Max),
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

fn number(value: Option<Scalar<'_>>) -> Option<f64> {
    match value {
        Some(Scalar::Number(value)) => Some(value),
        Some(Scalar::Int(value)) => Some(value as f64),
        _ => None,
    }
}

fn int(value: Option<Scalar<'_>>) -> u32 {
    match value {
        Some(Scalar::Int(value)) => u32::try_from(value).unwrap_or(0),
        _ => 0,
    }
}

/// The value of the chart write table's entry at `index`: a chart key, the
/// anchor's cell geometry, or a formatting key read back from the chart.
pub(crate) fn chart_value_at(
    summary: &Summary,
    cells: Option<[u32; 4]>,
    index: usize,
) -> Option<Value<'static>> {
    if let Some(key) = ChartKey::at(index) {
        return chart_value(summary, key);
    }
    match index {
        10..=13 => {
            let cells = cells?;
            Some(Value::Int(i64::from(cells[index - 10])))
        },
        _ => {
            let format = index.checked_sub(CHART_FORMAT_START)?;
            let (key, _) = shared::FORMAT_PROPS.get(format)?;
            summary.format(*key)
        },
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
    let series = item.summary().ok()?.series.get(item.child as usize)?;
    let key = [
        SeriesKey::Name,
        SeriesKey::Values,
        SeriesKey::Categories,
        SeriesKey::Color,
    ]
    .get(index)?;
    match key {
        SeriesKey::Name => series
            .name
            .as_deref()
            .map(|name| Value::Text(Cow::Borrowed(name))),
        SeriesKey::Values => Some(Value::Text(Cow::Owned(shared::join_values(&series.values)))),
        SeriesKey::Categories => (!series.categories.is_empty())
            .then(|| Value::Text(Cow::Owned(series.categories.join(",")))),
        SeriesKey::Color => series
            .color
            .as_deref()
            .map(|color| Value::Text(Cow::Borrowed(color))),
    }
}

/// The value of an axis node's key (`role`, `title`).
pub(crate) fn axis_value(item: &ChartItem, index: usize) -> Option<Value<'_>> {
    let axis = item.summary().ok()?.axes.get(item.child as usize)?;
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
    pub(super) fn sheet_chart(
        &self,
        sheet: usize,
        index: usize,
    ) -> Result<Option<Arc<shared::LazySummary>>> {
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
        Ok(xml.map(|xml| Arc::new(shared::LazySummary::new(xml))))
    }

    /// The two-cell anchor of the chart at `index` on the worksheet at
    /// `sheet`, as OfficeCLI's cell geometry `[x, y, width, height]`; read
    /// only when a view or a geometry `set` needs it.
    pub(super) fn sheet_chart_cells(&self, sheet: usize, index: usize) -> Result<Option<[u32; 4]>> {
        let anchor = self.0.owner.finish(|| {
            Ok(match &self.0.owner {
                super::XlsxOwner::Source(workbook) => match workbook.sheet(sheet)? {
                    Some(worksheet) if worksheet.kind() == xlsx::WorksheetKind::Worksheet => {
                        worksheet.chart_anchor(index)?
                    },
                    _ => None,
                },
                super::XlsxOwner::Eager(package) => match package.workbook_ref().sheet(sheet)? {
                    Some(worksheet) if worksheet.kind() == xlsx::WorksheetKind::Worksheet => {
                        package.chart_anchor(sheet, index)?
                    },
                    _ => None,
                },
            })
        })?;
        Ok(anchor.map(|anchor| {
            [
                anchor.from_col,
                anchor.from_row,
                anchor.to_col.saturating_sub(anchor.from_col),
                anchor.to_row.saturating_sub(anchor.from_row),
            ]
        }))
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
        let mut assigned = shared::Assigned::default();
        // `series{N}.values` given as a cell range, on any sheet (record
        // 0338): the series number and the range's absolute reference.
        let mut references: Vec<(usize, String, &'static str)> = Vec::new();
        let mut numbers: Vec<usize> = Vec::new();
        for (key, value) in node.props().iter() {
            let (index, entry) = CHART_WRITE_TABLE
                .resolve(key.as_str())
                .map_err(|e| e.with_property(key.as_str().to_owned()))?;
            if !assigned.insert(index) {
                return Err(error(
                    Code::InvalidArgument,
                    "a chart property was assigned more than once",
                )
                .with_property(entry.spec.name));
            }
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
                    ChartSetter::Numbered(number) => {
                        numbers.push(usize::from(number) + 1);
                        props.numbered_series(
                            usize::from(number) + 1,
                            &text(scalar).unwrap_or_default(),
                        )?;
                    },
                    ChartSetter::Dotted(number, field) => {
                        let number = usize::from(number) + 1;
                        let value = text(scalar).unwrap_or_default();
                        if field != shared::SeriesField::Color {
                            numbers.push(number);
                        }
                        if field == shared::SeriesField::Values && is_range(&value) {
                            let (values, reference) =
                                self.series_range(sheet, &value, entry.spec.name)?;
                            references.push((number, reference, entry.spec.name));
                            props.dotted_series(number, field, &values)?;
                        } else {
                            props.dotted_series(number, field, &value)?;
                        }
                    },
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
        if !references.is_empty() {
            let kind = props.variant.map_or(
                litchi_drawingml::chart::authoring::Kind::Column,
                |variant| variant.kind,
            );
            if binding.is_some()
                || !matches!(
                    kind,
                    litchi_drawingml::chart::authoring::Kind::Column
                        | litchi_drawingml::chart::authoring::Kind::Bar
                        | litchi_drawingml::chart::authoring::Kind::Line
                        | litchi_drawingml::chart::authoring::Kind::Area
                        | litchi_drawingml::chart::authoring::Kind::Pie
                        | litchi_drawingml::chart::authoring::Kind::Doughnut
                        | litchi_drawingml::chart::authoring::Kind::Radar
                        | litchi_drawingml::chart::authoring::Kind::Combo
                )
            {
                return Err(error(
                    Code::UnsupportedCapability,
                    "a series bound to a cell range is authored for category charts without dataRange",
                )
                .with_property(references[0].2)
                .with_suggestion("give the values inline (1,2,3), or bind a worksheet range with dataRange"));
            }
        }
        numbers.sort_unstable();
        numbers.dedup();
        let mut spec = props.into_spec()?;
        for (number, reference, _) in references {
            let index = numbers.iter().position(|held| *held == number).unwrap_or(0);
            spec = spec.with_series_reference(index, reference, None);
        }
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
        let source = self.editor_for_write()?;
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

    /// Read the values of a `series{N}.values` cell range (`Sheet1!B2:B13`,
    /// `'My Data'!$C$1:$C$4` or `B2:B5` on the chart's sheet; record 0338):
    /// the values as inline text the shared admission parses, and the
    /// range's absolute reference the chart refers to. The range is one row
    /// or one column; a cell without a number reads as 0, as `dataRange`.
    fn series_range(
        &self,
        sheet: usize,
        range: &str,
        key: &'static str,
    ) -> Result<(String, String)> {
        let invalid = |message: &'static str| {
            Error::new(Code::InvalidValue, message)
                .with_property(key)
                .with_format(FileFormat::Xlsx)
        };
        let range = range.trim();
        let (sheet_name, cells) = match range.rsplit_once('!') {
            Some((name, cells)) => {
                let name = name.trim();
                let name = match name
                    .strip_prefix('\'')
                    .and_then(|name| name.strip_suffix('\''))
                {
                    Some(quoted) => quoted.replace("''", "'"),
                    None => name.to_owned(),
                };
                (name, cells)
            },
            None => (super::node::sheet_name(self, sheet).to_owned(), range),
        };
        let rect = cells
            .replace('$', "")
            .parse::<xlsx::Rect>()
            .map_err(|_| invalid("the series range is not an A1 range"))?;
        let (first_row, first_column) = (rect.start().row().get(), rect.start().column().get());
        let (last_row, last_column) = (
            rect.end().0.saturating_sub(1),
            rect.end().1.saturating_sub(1),
        );
        if first_row != last_row && first_column != last_column {
            return Err(invalid("the series range is one row or one column"));
        }
        let package = self.editor_for_write()?;
        let workbook = package.workbook();
        let worksheet = workbook
            .sheet(sheet_name.as_str())?
            .ok_or_else(|| invalid("the series range names no worksheet"))?;
        let mut values = Vec::new();
        for row in first_row..=last_row {
            for column in first_column..=last_column {
                let address = xlsx::Address::at(row, column)
                    .map_err(|_| invalid("the series range is outside the worksheet"))?;
                let number = match worksheet.cell(address)?.stored() {
                    Some(xlsx::Cell::Value(value)) => cell_value(value).1,
                    Some(xlsx::Cell::Formula(formula)) => formula
                        .cached()
                        .and_then(|cache| cell_value(cache.value()).1),
                    _ => None,
                };
                values.push(number.unwrap_or(0.0).to_string());
            }
        }
        let column = |index: u32| litchi_drawingml::chart::authoring::column_name(index as usize);
        let reference = if first_row == last_row && first_column == last_column {
            format!(
                "{}!${}${}",
                quote_sheet(worksheet.name()),
                column(first_column),
                first_row + 1
            )
        } else {
            format!(
                "{}!${}${}:${}${}",
                quote_sheet(worksheet.name()),
                column(first_column),
                first_row + 1,
                column(last_column),
                last_row + 1
            )
        };
        Ok((values.join(","), reference))
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
        let package = self.editor_for_write()?;
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
        let mut geometry = [None::<u32>; 4];
        let summary = || item.summary();
        for (key, value) in props.iter() {
            match node.item() {
                Item::Chart(_) => {
                    let (_, entry) = CHART_WRITE_TABLE
                        .resolve(key.as_str())
                        .map_err(|e| e.with_property(key.as_str().to_owned()))?;
                    let name = entry.spec.name;
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
                            ChartSetter::CatTitle | ChartSetter::AxisTitle => {
                                let role = if matches!(setter, ChartSetter::CatTitle) {
                                    Role::Category
                                } else {
                                    Role::Value
                                };
                                let summary = summary()?;
                                let axis = summary.axis(role).ok_or_else(|| {
                                    error(Code::UnsupportedCapability, "the chart has no such axis")
                                        .with_property(name)
                                })?;
                                edit.axis_titles.push((axis.id, text(scalar)));
                            },
                            ChartSetter::Dotted(number, field) => edit.dotted_series(
                                summary()?,
                                usize::from(number) + 1,
                                field,
                                text(scalar),
                            )?,
                            ChartSetter::ChartType => {
                                edit.chart_type(summary()?, &text(scalar).unwrap_or_default())?;
                            },
                            ChartSetter::Data => {
                                edit.data(summary()?, &text(scalar).unwrap_or_default())?;
                            },
                            ChartSetter::Format(key) => {
                                edit.format(summary()?, key, name, scalar)?
                            },
                            ChartSetter::X => geometry[0] = Some(int(scalar)),
                            ChartSetter::Y => geometry[1] = Some(int(scalar)),
                            ChartSetter::Width => geometry[2] = Some(int(scalar)),
                            ChartSetter::Height => geometry[3] = Some(int(scalar)),
                            ChartSetter::DataRange
                            | ChartSetter::Name
                            | ChartSetter::Numbered(_) => {
                                unreachable!("chart set admission")
                            },
                        }
                        Ok(())
                    })?;
                },
                Item::Series(_) => {
                    let index = item
                        .summary()?
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
                            SeriesSetter::Color => edit.series_colors.push((
                                index,
                                text(scalar)
                                    .map(|color| shared::parse_color(&color, "color"))
                                    .transpose()?,
                            )),
                        }
                        Ok(())
                    })?;
                },
                _ => {
                    let axis = item
                        .summary()?
                        .axes
                        .get(item.child as usize)
                        .ok_or_else(|| error(Code::PathNotFound, "the axis is missing"))?;
                    let (_, entry) = AXIS_WRITE_TABLE
                        .resolve(key.as_str())
                        .map_err(|e| e.with_property(key.as_str().to_owned()))?;
                    entry.prepare(Operation::Set, value, |setter, scalar| {
                        match setter {
                            AxisSetter::Title => edit.axis_titles.push((axis.id, text(scalar))),
                            AxisSetter::Min | AxisSetter::Max => edit.axis_bound(
                                axis,
                                matches!(setter, AxisSetter::Max),
                                number(scalar),
                            )?,
                        }
                        Ok(())
                    })?;
                },
            }
        }
        let sheet = node.sheet_position();
        let source = self.editor_for_write()?;
        let mut native = source.edit_charts()?;
        if geometry.iter().any(Option::is_some) {
            let current = self
                .sheet_chart_cells(sheet, item.chart as usize)?
                .ok_or_else(|| {
                    error(
                        Code::UnsupportedCapability,
                        "only a chart placed by a two-cell anchor can be moved",
                    )
                    .with_property("x")
                })?;
            let [x, y, width, height] =
                [0, 1, 2, 3].map(|slot| geometry[slot].unwrap_or(current[slot]));
            let (to_col, to_row) = (x.saturating_add(width), y.saturating_add(height));
            if to_col >= 16_384 || to_row >= 1_048_576 {
                return Err(
                    error(Code::OutOfRange, "the chart extends past the worksheet")
                        .with_property("width"),
                );
            }
            native.set_anchor(
                sheet,
                item.chart as usize,
                &xlsx::chart::Anchor::new(x, y, to_col, to_row),
            )?;
        }
        let xml = native.chart(sheet, item.chart as usize)?;
        if let Some(edited) = edit.apply(xml, false)? {
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
        let source = self.editor_for_write()?;
        let mut native = source.edit_charts()?;
        native.remove(sheet, item.chart as usize)?;
        let target = native.commit()?.into_package();
        let after = self.adopt_workbook(target.clone());
        Ok(Some(Commit::from_metadata(after, source, target)))
    }
}

/// Whether a series value is a cell range (`Sheet1!B2:B5`, `B2:B5`), as
/// the shared admission tells one from inline values.
fn is_range(text: &str) -> bool {
    let text = text.trim();
    if text.contains('!') {
        return true;
    }
    let mut cells = text.split(':');
    let cell = |part: Option<&str>| {
        part.is_some_and(|part| {
            let part = part.trim().replace('$', "");
            let letters = part.bytes().take_while(u8::is_ascii_alphabetic).count();
            letters > 0
                && part.len() > letters
                && part.bytes().skip(letters).all(|byte| byte.is_ascii_digit())
        })
    };
    cell(cells.next()) && cell(cells.next()) && cells.next().is_none()
}

/// A sheet name in a formula reference: quoted unless it is a plain name.
fn quote_sheet(name: &str) -> String {
    let plain = name
        .chars()
        .next()
        .is_some_and(|first| first.is_alphabetic() || first == '_')
        && name
            .chars()
            .all(|c| c.is_alphanumeric() || c == '_' || c == '.');
    if plain {
        name.to_owned()
    } else {
        format!("'{}'", name.replace('\'', "''"))
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
