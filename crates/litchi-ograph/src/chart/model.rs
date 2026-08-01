//! Host-neutral semantic chart values.

use super::{Kind, Ref, Stream, axis, codec, format, group};
use crate::{Error, Limits, Result};

/// Number of points in one BIFF chart series (`0..=32_767`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
pub struct Count(u16);

impl Count {
    /// Zero points.
    pub const ZERO: Self = Self(0);

    /// Creates a checked BIFF chart count.
    pub const fn new(value: u16) -> Option<Self> {
        if value <= 32_767 {
            Some(Self(value))
        } else {
            None
        }
    }

    /// Returns the stored count.
    pub const fn get(self) -> u16 {
        self.0
    }
}

/// Chart-group index or order (`0..=9`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
pub struct GroupId(u8);

impl GroupId {
    /// The primary chart group.
    pub const ZERO: Self = Self(0);

    /// Creates a checked BIFF chart-group identifier.
    pub const fn new(value: u8) -> Option<Self> {
        if value <= 9 { Some(Self(value)) } else { None }
    }

    /// Returns the stored identifier.
    pub const fn get(self) -> u8 {
        self.0
    }
}

/// Producer context required to interpret and emit a chart substream.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Context {
    kind: Kind,
    external_sheets: Option<usize>,
}

impl Context {
    /// Excel-hosted chart context.
    pub const fn excel() -> Self {
        Self {
            kind: Kind::Excel,
            external_sheets: None,
        }
    }

    /// Standalone Microsoft Graph chart context.
    pub const fn graph() -> Self {
        Self {
            kind: Kind::Graph,
            external_sheets: None,
        }
    }

    /// Adds the known number of entries in the host's external-sheet table.
    #[must_use]
    pub const fn with_external_sheets(mut self, count: usize) -> Self {
        self.external_sheets = Some(count);
        self
    }

    /// Chart BOF grammar selected by this context.
    pub const fn kind(self) -> Kind {
        self.kind
    }

    /// Known external-sheet table size, when supplied by the host.
    pub const fn external_sheet_count(self) -> Option<usize> {
        self.external_sheets
    }
}

impl From<Kind> for Context {
    fn from(kind: Kind) -> Self {
        match kind {
            Kind::Graph => Self::graph(),
            Kind::Excel => Self::excel(),
        }
    }
}

/// Chart rectangle in the host's BIFF coordinate units.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Rect {
    pub x: i32,
    pub y: i32,
    pub width: i32,
    pub height: i32,
}

impl Default for Rect {
    fn default() -> Self {
        Self {
            x: 0,
            y: 0,
            width: 4_000 << 16,
            height: 3_000 << 16,
        }
    }
}

/// Sheet-level chart properties.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Props {
    /// Raw defined ShtProps flags and blank-display mode.
    pub flags: u32,
    /// Whether a PlotArea record is present.
    pub plot_area: bool,
}

impl Default for Props {
    fn default() -> Self {
        Self {
            flags: 2,
            plot_area: true,
        }
    }
}

/// Cell value kind used by a chart series.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DataKind {
    Numeric,
    Text,
}

/// Semantic role of a series data link.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum Role {
    Name = 0,
    Values = 1,
    Categories = 2,
    Bubbles = 3,
}

/// Source selected by a series data link.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum Source {
    Automatic = 0,
    Literal = 1,
    Cells = 2,
}

/// Datasheet row or column used by a standalone Graph BRAI (`0..=3_999`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
pub struct RowCol(u16);

impl RowCol {
    pub const ZERO: Self = Self(0);

    pub const fn new(value: u16) -> Option<Self> {
        if value <= 0x0F9F {
            Some(Self(value))
        } else {
            None
        }
    }

    pub const fn get(self) -> u16 {
        self.0
    }
}

/// One checked BIFF8 cell or rectangular cell-range reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CellRef {
    pub external_sheet: u16,
    pub first_row: u16,
    pub last_row: u16,
    pub first_col: u8,
    pub last_col: u8,
}

/// Inert, producer-specific data link. Formula tokens are never evaluated.
#[derive(Debug, PartialEq, Eq)]
pub enum Link {
    /// Fixed-size `[MS-OGRAPH]` BRAI using a datasheet row or column.
    Graph {
        role: Role,
        source: Source,
        unlinked_format: bool,
        number_format: u16,
        row_col: RowCol,
    },
    /// Variable-size `[MS-XLS]` BRAI using a ChartParsedFormula.
    Excel {
        role: Role,
        source: Source,
        unlinked_format: bool,
        number_format: u16,
        formula: Vec<u8>,
        refs: Vec<CellRef>,
    },
}

impl Link {
    /// Creates a standalone Graph link.
    pub const fn graph(role: Role, source: Source, row_col: RowCol) -> Self {
        Self::Graph {
            role,
            source,
            unlinked_format: false,
            number_format: 0,
            row_col,
        }
    }

    /// Creates an Excel link, moving its inert formula token allocation.
    pub const fn excel(role: Role, source: Source, formula: Vec<u8>) -> Self {
        Self::Excel {
            role,
            source,
            unlinked_format: false,
            number_format: 0,
            formula,
            refs: Vec::new(),
        }
    }

    pub const fn role(&self) -> Role {
        match self {
            Self::Graph { role, .. } | Self::Excel { role, .. } => *role,
        }
    }

    pub const fn source(&self) -> Source {
        match self {
            Self::Graph { source, .. } | Self::Excel { source, .. } => *source,
        }
    }
}

/// One chart series.
#[derive(Debug, PartialEq, Eq)]
pub struct Series {
    pub category_kind: DataKind,
    pub category_count: Count,
    pub value_count: Count,
    pub bubble_count: Count,
    pub group: GroupId,
    pub name: Option<String>,
    pub links: Vec<Link>,
}

impl Series {
    /// Creates an empty text-category series in the primary chart group.
    pub const fn new() -> Self {
        Self {
            category_kind: DataKind::Text,
            category_count: Count::ZERO,
            value_count: Count::ZERO,
            bubble_count: Count::ZERO,
            group: GroupId::ZERO,
            name: None,
            links: Vec::new(),
        }
    }
}

impl Default for Series {
    fn default() -> Self {
        Self::new()
    }
}

/// Chart-family configuration attached to one group.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Family {
    Line {
        flags: u16,
    },
    Bar {
        overlap: group::Overlap,
        gap: group::Gap,
        flags: u16,
    },
    Area {
        flags: u16,
    },
    Pie {
        rotation: u16,
        hole: u16,
        flags: u16,
    },
    Scatter {
        bubble_percent: group::BubblePercent,
        bubble_kind: group::BubbleKind,
        flags: u16,
    },
    Radar {
        filled: bool,
        flags: u16,
    },
    Surface {
        flags: u16,
    },
}

/// One ordered chart group.
#[derive(Debug, PartialEq, Eq)]
pub struct Group {
    pub order: GroupId,
    pub vary_colors: bool,
    pub family: Family,
    pub lines: Vec<group::Line>,
    pub drop_bars: Vec<group::DropBar>,
}

impl Group {
    /// Primary line-chart group used by a new chart.
    pub const fn line() -> Self {
        Self {
            order: GroupId::ZERO,
            vary_colors: false,
            family: Family::Line { flags: 0 },
            lines: Vec::new(),
            drop_bars: Vec::new(),
        }
    }
}

/// One cached chart value.
#[derive(Debug, PartialEq)]
pub enum Value {
    Number(f64),
    Text(String),
    Blank,
}

/// Producer-specific cache cell coordinate.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Cell {
    /// Excel BIFF8 grid coordinate.
    Excel { row: u16, col: u8 },
    /// Standalone Graph datasheet coordinate (`0..=3_999` per dimension).
    Graph { row: RowCol, col: RowCol },
}

/// Cell-shaped cache entry associated with a cache index.
#[derive(Debug, PartialEq)]
pub struct Cache {
    pub index: u16,
    pub cell: Cell,
    pub format: u16,
    pub value: Value,
}

/// Legend rectangle and layout properties.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Legend {
    pub x: i32,
    pub y: i32,
    pub width: i32,
    pub height: i32,
    pub position: u8,
    pub spacing: u8,
    pub flags: u16,
}

/// Opaque data-label record retained as part of the semantic inventory.
#[derive(Debug, PartialEq, Eq)]
pub struct Label {
    pub kind: crate::raw::Kind,
    pub data: Vec<u8>,
}

/// Opaque record observed during parsing, in original record order.
#[derive(Debug, PartialEq, Eq)]
pub struct Raw {
    kind: crate::raw::Kind,
    data: Vec<u8>,
    offset: usize,
}

impl Raw {
    /// BIFF record identifier.
    pub const fn kind(&self) -> crate::raw::Kind {
        self.kind
    }

    /// Exact record payload.
    pub fn data(&self) -> &[u8] {
        &self.data
    }

    /// Original byte offset in the chart substream.
    pub const fn offset(&self) -> usize {
        self.offset
    }

    pub(super) fn parsed(kind: crate::raw::Kind, data: Vec<u8>, offset: usize) -> Self {
        Self { kind, data, offset }
    }
}

#[derive(Debug)]
pub(super) enum Origin {
    Fresh,
    Parsed(Stream),
}

/// Host-neutral semantic chart.
///
/// Parsed values retain their exact [`Stream`]. An untouched parsed value can
/// therefore be encoded byte-for-byte without copying. Mutation is allowed for
/// inspection workflows, but encoding such a value is refused until a future
/// lossless record editor can prove placement of every opaque record.
/// Fresh values can be assembled and validated, but encoding currently returns
/// [`Error::UnsupportedAuthoring`] until the complete mandatory chart-sheet
/// scaffold is represented. This prevents self-consistent but Office-invalid
/// streams from escaping the crate.
#[derive(Debug)]
pub struct Chart {
    pub(super) context: Context,
    pub(super) rect: Rect,
    pub(super) props: Props,
    pub(super) title: Option<String>,
    pub(super) series: Vec<Series>,
    pub(super) groups: Vec<Group>,
    pub(super) axes: Vec<axis::Axis>,
    pub(super) legend: Option<Legend>,
    pub(super) caches: Vec<Cache>,
    pub(super) formats: Vec<format::Format>,
    pub(super) labels: Vec<Label>,
    pub(super) unknown: Vec<Raw>,
    pub(super) origin: Origin,
    pub(super) dirty: bool,
    pub(super) limits: Limits,
    /// Internal proof gate. No public constructor enables it until the full
    /// CHARTSHEET/CHARTFOMATS/SERIESDATA grammar is modeled.
    pub(super) authoring_proven: bool,
}

impl Chart {
    /// Creates a fresh chart using conservative limits.
    pub fn new(context: Context) -> Result<Self> {
        Self::new_with(context, Limits::default())
    }

    /// Creates a fresh chart with explicit authoring bounds.
    pub fn new_with(context: Context, limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        let mut groups = Vec::new();
        groups.try_reserve_exact(1).map_err(|_| Error::Allocation {
            resource: "chart groups",
        })?;
        groups.push(Group::line());
        Ok(Self {
            context,
            rect: Rect::default(),
            props: Props::default(),
            title: None,
            series: Vec::new(),
            groups,
            axes: Vec::new(),
            legend: None,
            caches: Vec::new(),
            formats: Vec::new(),
            labels: Vec::new(),
            unknown: Vec::new(),
            origin: Origin::Fresh,
            dirty: false,
            limits,
            authoring_proven: false,
        })
    }

    /// Parses a borrowed chart and retains an exact bounded copy for replay.
    pub fn parse(input: Ref<'_>, context: Context) -> Result<Self> {
        let limits = input.limits();
        Self::parse_with(input, context, limits)
    }

    /// Parses a borrowed chart under explicit semantic limits.
    pub fn parse_with(input: Ref<'_>, context: Context, limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        let mut chart = codec::parse(input, context, limits)?;
        chart.origin = Origin::Parsed(input.own_with(limits)?);
        Ok(chart)
    }

    /// Parses a move-owned stream without copying its input allocation.
    pub fn open(input: Stream, context: Context) -> Result<Self> {
        let limits = input.as_ref().limits();
        Self::open_with(input, context, limits)
    }

    /// Parses a move-owned stream under explicit semantic limits.
    pub fn open_with(input: Stream, context: Context, limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        let input = input.relimit(limits)?;
        let mut chart = codec::parse(input.as_ref(), context, limits)?;
        chart.origin = Origin::Parsed(input);
        Ok(chart)
    }

    /// Replays an untouched parsed chart, or refuses unsupported fresh authoring.
    pub fn encode(self) -> Result<Stream> {
        let limits = self.limits;
        self.encode_with(limits)
    }

    /// Consumes and replays under explicit bounds.
    ///
    /// Parsed mutations return [`Error::UnsafeEdit`]; fresh values return
    /// [`Error::UnsupportedAuthoring`] until full chart-sheet authoring lands.
    pub fn encode_with(mut self, limits: Limits) -> Result<Stream> {
        let limits = limits.validate()?;
        let origin = std::mem::replace(&mut self.origin, Origin::Fresh);
        match origin {
            Origin::Parsed(_stream) if self.dirty => Err(Error::UnsafeEdit {
                reason: "opaque or reserved source records could not be placed losslessly",
            }),
            Origin::Parsed(stream) => stream.relimit(limits),
            Origin::Fresh => {
                let bytes = codec::encode(&self, limits)?;
                Stream::with_limits(bytes, limits)
            },
        }
    }

    pub const fn context(&self) -> Context {
        self.context
    }

    pub const fn rect(&self) -> Rect {
        self.rect
    }

    pub fn set_rect(&mut self, value: Rect) {
        self.touch();
        self.rect = value;
    }

    pub const fn props(&self) -> Props {
        self.props
    }

    pub fn set_props(&mut self, value: Props) {
        self.touch();
        self.props = value;
    }

    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    pub fn set_title(&mut self, value: Option<String>) {
        self.touch();
        self.title = value;
    }

    pub fn series(&self) -> &[Series] {
        &self.series
    }

    pub fn series_mut(&mut self) -> &mut [Series] {
        self.touch();
        &mut self.series
    }

    pub fn add_series(&mut self, value: Series) -> Result<()> {
        check_add(self.series.len(), self.limits.max_series, "series count")?;
        reserve_one(&mut self.series, "chart series")?;
        self.touch();
        self.series.push(value);
        Ok(())
    }

    pub fn remove_series(&mut self, index: usize) -> Option<Series> {
        if index >= self.series.len() {
            return None;
        }
        self.touch();
        Some(self.series.remove(index))
    }

    pub fn groups(&self) -> &[Group] {
        &self.groups
    }

    pub fn groups_mut(&mut self) -> &mut [Group] {
        self.touch();
        &mut self.groups
    }

    pub fn add_group(&mut self, value: Group) -> Result<()> {
        check_add(self.groups.len(), self.limits.max_groups, "group count")?;
        reserve_one(&mut self.groups, "chart groups")?;
        self.touch();
        self.groups.push(value);
        Ok(())
    }

    pub fn remove_group(&mut self, index: usize) -> Option<Group> {
        if index >= self.groups.len() {
            return None;
        }
        self.touch();
        Some(self.groups.remove(index))
    }

    pub fn axes(&self) -> &[axis::Axis] {
        &self.axes
    }

    pub fn axes_mut(&mut self) -> &mut [axis::Axis] {
        self.touch();
        &mut self.axes
    }

    pub fn add_axis(&mut self, value: axis::Axis) -> Result<()> {
        check_add(self.axes.len(), self.limits.max_axes, "axis count")?;
        reserve_one(&mut self.axes, "chart axes")?;
        self.touch();
        self.axes.push(value);
        Ok(())
    }

    pub fn remove_axis(&mut self, index: usize) -> Option<axis::Axis> {
        if index >= self.axes.len() {
            return None;
        }
        self.touch();
        Some(self.axes.remove(index))
    }

    pub const fn legend(&self) -> Option<Legend> {
        self.legend
    }

    pub fn set_legend(&mut self, value: Option<Legend>) {
        self.touch();
        self.legend = value;
    }

    pub fn caches(&self) -> &[Cache] {
        &self.caches
    }

    pub fn caches_mut(&mut self) -> &mut [Cache] {
        self.touch();
        &mut self.caches
    }

    pub fn add_cache(&mut self, value: Cache) -> Result<()> {
        check_add(
            self.caches.len(),
            self.limits.max_cached_values,
            "cached value count",
        )?;
        reserve_one(&mut self.caches, "chart cache")?;
        self.touch();
        self.caches.push(value);
        Ok(())
    }

    pub fn formats(&self) -> &[format::Format] {
        &self.formats
    }

    pub fn formats_mut(&mut self) -> &mut [format::Format] {
        self.touch();
        &mut self.formats
    }

    pub fn add_format(&mut self, value: format::Format) -> Result<()> {
        check_add(
            self.formats.len(),
            self.limits.max_chart_records,
            "chart record count",
        )?;
        reserve_one(&mut self.formats, "chart formats")?;
        self.touch();
        self.formats.push(value);
        Ok(())
    }

    pub fn labels(&self) -> &[Label] {
        &self.labels
    }

    pub fn labels_mut(&mut self) -> &mut [Label] {
        self.touch();
        &mut self.labels
    }

    pub fn add_label(&mut self, value: Label) -> Result<()> {
        check_add(
            self.labels.len(),
            self.limits.max_chart_records,
            "chart record count",
        )?;
        reserve_one(&mut self.labels, "chart labels")?;
        self.touch();
        self.labels.push(value);
        Ok(())
    }

    /// Unknown and recognized-but-opaque records in original encounter order.
    pub fn unknown(&self) -> &[Raw] {
        &self.unknown
    }

    /// Whether this parsed chart still has an exact replayable source stream.
    pub fn is_pristine(&self) -> bool {
        matches!(&self.origin, Origin::Parsed(_)) && !self.dirty
    }

    fn touch(&mut self) {
        if matches!(&self.origin, Origin::Parsed(_)) {
            self.dirty = true;
        }
    }
}

fn reserve_one<T>(values: &mut Vec<T>, resource: &'static str) -> Result<()> {
    values
        .try_reserve(1)
        .map_err(|_| Error::Allocation { resource })
}

fn check_add(current: usize, maximum: usize, resource: &'static str) -> Result<()> {
    let observed = current
        .checked_add(1)
        .ok_or(Error::SizeOverflow { resource })?;
    if observed > maximum {
        return Err(Error::LimitExceeded {
            resource,
            observed: crate::limits::as_u64(observed),
            maximum: crate::limits::as_u64(maximum),
        });
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::chart::{self, RowCol};
    use crate::raw::{Encoder, Kind as RecordKind, Records};

    const UNKNOWN: RecordKind = RecordKind::new(0x7777);

    fn count(value: u16) -> Count {
        Count::new(value).expect("bounded fixture count")
    }

    fn line_format() -> chart::format::Line {
        chart::format::Line {
            color: [1, 2, 3, 0],
            pattern: 0,
            weight: 0,
            flags: 0,
            color_index: 8,
        }
    }

    fn area_format() -> chart::format::Area {
        chart::format::Area {
            foreground: [4, 5, 6, 0],
            background: [7, 8, 9, 0],
            pattern: 1,
            flags: 0,
            foreground_index: 9,
            background_index: 10,
        }
    }

    fn fixture(mut chart: Chart) -> chart::Stream {
        chart.authoring_proven = true;
        chart.encode().expect("internal parser fixture")
    }

    fn excel_chart() -> Chart {
        let mut chart = Chart::new(Context::excel().with_external_sheets(1)).expect("new chart");
        chart.set_title(Some("Revenue".into()));
        chart
            .add_series(Series {
                category_kind: DataKind::Text,
                category_count: count(2),
                value_count: count(2),
                bubble_count: Count::ZERO,
                group: GroupId::ZERO,
                name: Some("FY26".into()),
                links: vec![Link::excel(
                    Role::Values,
                    Source::Cells,
                    vec![0x1B, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0],
                )],
            })
            .expect("series");
        chart
            .add_cache(Cache {
                index: 1,
                cell: Cell::Excel { row: 0, col: 0 },
                format: 0,
                value: Value::Number(42.5),
            })
            .expect("numeric cache");
        chart
            .add_cache(Cache {
                index: 1,
                cell: Cell::Excel { row: 1, col: 0 },
                format: 2,
                value: Value::Text("safe".into()),
            })
            .expect("text cache");
        chart
            .add_cache(Cache {
                index: 1,
                cell: Cell::Excel { row: 2, col: 0 },
                format: 3,
                value: Value::Blank,
            })
            .expect("blank cache");
        chart
    }

    #[test]
    fn fresh_encode_refuses_while_internal_fixture_replay_moves_original_stream() {
        assert!(matches!(
            excel_chart().encode(),
            Err(Error::UnsupportedAuthoring { .. })
        ));
        let stream = fixture(excel_chart());
        let pointer = stream.as_bytes().as_ptr();
        let parsed =
            Chart::open(stream, Context::excel().with_external_sheets(1)).expect("semantic parse");
        assert!(parsed.is_pristine());
        assert_eq!(parsed.title(), Some("Revenue"));
        assert_eq!(parsed.series().len(), 1);
        assert_eq!(parsed.caches().len(), 3);
        assert!(matches!(parsed.caches()[2].value, Value::Blank));
        let replay = parsed.encode().expect("exact replay");
        assert_eq!(replay.as_bytes().as_ptr(), pointer);
    }

    #[test]
    fn graph_and_excel_links_have_distinct_checked_wire_grammars() {
        let row_col = RowCol::new(7).expect("Graph coordinate");
        let mut graph = Chart::new(Context::graph()).expect("Graph chart");
        graph
            .add_series(Series {
                links: vec![Link::graph(Role::Values, Source::Literal, row_col)],
                ..Series::new()
            })
            .expect("Graph series");
        graph
            .add_cache(Cache {
                index: 0,
                cell: Cell::Graph {
                    row: RowCol::new(1).expect("row"),
                    col: RowCol::new(2).expect("column"),
                },
                format: 4,
                value: Value::Blank,
            })
            .expect("Graph blank");
        let stream = fixture(graph);
        let parsed = Chart::open(stream, Context::graph()).expect("Graph parse");
        assert!(matches!(parsed.series()[0].links[0], Link::Graph { .. }));
        assert!(matches!(parsed.caches()[0].cell, Cell::Graph { .. }));

        let mut wrong = Chart::new(Context::graph()).expect("Graph chart");
        wrong
            .add_series(Series {
                links: vec![Link::excel(Role::Values, Source::Automatic, Vec::new())],
                ..Series::new()
            })
            .expect("series");
        wrong.authoring_proven = true;
        assert!(matches!(
            wrong.encode(),
            Err(Error::InvalidModel { field: "link", .. })
        ));
    }

    #[test]
    fn parsed_mutation_is_refused_and_unknown_order_replays_exactly() {
        let original = fixture(excel_chart());
        let mut out = Encoder::new();
        for item in Records::new(original.as_bytes()) {
            let record = item.expect("valid record");
            if record.kind() == super::super::EOF {
                out.push(UNKNOWN, &[9, 8, 7]).expect("unknown record");
            }
            out.push_ref(record).expect("record replay");
        }
        let bytes = out.finish();
        let pointer = bytes.as_ptr();
        let stream = chart::Stream::open(bytes).expect("raw chart");
        let parsed =
            Chart::open(stream, Context::excel().with_external_sheets(1)).expect("semantic chart");
        assert_eq!(parsed.unknown().len(), 1);
        assert_eq!(parsed.unknown()[0].kind(), UNKNOWN);
        let replay = parsed.encode().expect("exact replay");
        assert_eq!(replay.as_bytes().as_ptr(), pointer);
        let mut parsed =
            Chart::open(replay, Context::excel().with_external_sheets(1)).expect("semantic chart");
        parsed.set_title(Some("Changed".into()));
        assert!(matches!(parsed.encode(), Err(Error::UnsafeEdit { .. })));
    }

    #[test]
    fn rejects_context_mismatch_tighter_replay_limit_and_invalid_properties() {
        let stream = fixture(excel_chart());
        let bytes = stream.as_bytes().len();
        assert!(matches!(
            Chart::open(stream, Context::graph()),
            Err(Error::InvalidModel {
                field: "context",
                ..
            })
        ));

        let stream = fixture(excel_chart());
        let chart =
            Chart::open(stream, Context::excel().with_external_sheets(1)).expect("parsed chart");
        assert!(matches!(
            chart.encode_with(Limits {
                max_output_bytes: bytes.saturating_sub(1),
                ..Limits::default()
            }),
            Err(Error::LimitExceeded {
                resource: "output bytes",
                ..
            })
        ));

        let mut valid_blank_mode = Chart::new(Context::excel()).expect("chart");
        valid_blank_mode.set_props(Props {
            flags: 2 | (2 << 16),
            plot_area: true,
        });
        valid_blank_mode.authoring_proven = true;
        assert!(valid_blank_mode.encode().is_ok());

        let mut reserved = Chart::new(Context::excel()).expect("chart");
        reserved.set_props(Props {
            flags: 1 << 2,
            plot_area: true,
        });
        reserved.authoring_proven = true;
        assert!(matches!(reserved.encode(), Err(Error::InvalidModel { .. })));

        let mut dependency = Chart::new(Context::excel()).expect("chart");
        dependency.set_props(Props {
            flags: 1 << 4,
            plot_area: true,
        });
        dependency.authoring_proven = true;
        assert!(matches!(
            dependency.encode(),
            Err(Error::InvalidModel { .. })
        ));
    }

    #[test]
    fn add_methods_enforce_authoring_limits_before_growth() {
        let limits = Limits {
            max_series: 1,
            max_groups: 1,
            max_axes: 1,
            max_cached_values: 1,
            ..Limits::default()
        };
        let mut chart = Chart::new_with(Context::excel(), limits).expect("bounded chart");
        chart.add_series(Series::new()).expect("first series");
        assert!(matches!(
            chart.add_series(Series::new()),
            Err(Error::LimitExceeded {
                resource: "series count",
                ..
            })
        ));
        assert!(matches!(
            chart.add_group(Group::line()),
            Err(Error::LimitExceeded {
                resource: "group count",
                ..
            })
        ));
        chart
            .add_axis(axis::Axis::new(axis::Kind::Category))
            .expect("first axis");
        assert!(chart.add_axis(axis::Axis::new(axis::Kind::Value)).is_err());
    }

    #[test]
    fn group_lines_and_drop_bars_emit_mandatory_owned_formats() {
        let mut chart = Chart::new(Context::excel()).expect("chart");
        let group = chart.groups_mut().first_mut().expect("default line group");
        group
            .lines
            .try_reserve_exact(1)
            .expect("line fixture allocation");
        group.lines.push(chart::group::Line {
            kind: crate::record::line::Kind::HighLow,
            format: line_format(),
        });
        group
            .drop_bars
            .try_reserve_exact(1)
            .expect("DropBar fixture allocation");
        group.drop_bars.push(chart::group::DropBar {
            gap: chart::group::Gap::new(20).expect("bounded gap"),
            line: line_format(),
            area: area_format(),
        });

        let stream = fixture(chart);
        let kinds = stream
            .records()
            .map(|record| record.expect("valid record").kind())
            .collect::<Vec<_>>();
        let crt = kinds
            .iter()
            .position(|kind| *kind == RecordKind::new(0x101C))
            .expect("CrtLine");
        assert_eq!(kinds.get(crt + 1), Some(&RecordKind::new(0x1007)));
        let drop = kinds
            .iter()
            .position(|kind| *kind == RecordKind::new(0x103D))
            .expect("DropBar");
        assert_eq!(
            kinds.get(drop..drop + 5),
            Some(
                [0x103D, 0x1033, 0x1007, 0x100A, 0x1034]
                    .map(RecordKind::new)
                    .as_slice()
            )
        );

        let parsed = Chart::open(stream, Context::excel()).expect("parse");
        let group = parsed.groups().first().expect("group");
        assert_eq!(group.lines.len(), 1);
        assert_eq!(group.drop_bars.len(), 1);
        assert_eq!(group.drop_bars[0].gap.get(), 20);
    }

    #[test]
    fn rejects_missing_collection_begin_and_nesting_over_limit() {
        let stream = fixture(excel_chart());
        let mut out = Encoder::new();
        let mut after_series = false;
        let mut removed = false;
        for item in Records::new(stream.as_bytes()) {
            let record = item.expect("valid record");
            if after_series && record.kind() == RecordKind::new(0x1033) {
                removed = true;
                after_series = false;
                continue;
            }
            after_series = record.kind() == RecordKind::new(0x1003);
            out.push_ref(record).expect("record replay");
        }
        assert!(removed);
        let malformed = out.finish();
        let input = chart::Ref::open(&malformed).expect("raw boundaries remain valid");
        assert!(matches!(
            Chart::parse(input, Context::excel().with_external_sheets(1)),
            Err(Error::InvalidChart {
                reason: "collection-owning record is not followed immediately by Begin",
                ..
            })
        ));

        assert!(matches!(
            Chart::parse_with(
                stream.as_ref(),
                Context::excel().with_external_sheets(1),
                Limits {
                    max_nesting: 2,
                    ..Limits::default()
                }
            ),
            Err(Error::LimitExceeded {
                resource: "chart nesting",
                ..
            })
        ));
    }
}
