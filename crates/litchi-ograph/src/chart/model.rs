//! Host-neutral semantic chart values.

use std::ops::{Deref, DerefMut};

use super::{Kind, Ref, Stream, axis, cache, codec, format, group, layout};
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

/// Zero-based chart-group index used by `SerToCrt` (`0..=9`).
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

/// Chart-group drawing order used by `ChartFormat.icrt` (`0..=9`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
pub struct Order(u8);

impl Order {
    /// Bottom of the chart-group z-order.
    pub const ZERO: Self = Self(0);

    /// Creates a checked drawing order.
    pub const fn new(value: u8) -> Option<Self> {
        if value <= 9 { Some(Self(value)) } else { None }
    }

    /// Returns the stored drawing order.
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

impl Role {
    /// Mandatory regular-series AI order.
    pub const ALL: [Self; 4] = [Self::Name, Self::Values, Self::Categories, Self::Bubbles];
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

/// One mandatory AI binding: a producer-specific BRAI and its optional
/// immediately following `SeriesText`.
#[derive(Debug, PartialEq, Eq)]
pub struct Binding {
    link: Link,
    text: Option<String>,
}

impl Binding {
    /// Creates a binding by moving its inert link and optional text.
    pub const fn new(link: Link, text: Option<String>) -> Self {
        Self { link, text }
    }

    /// Producer-specific data link.
    pub const fn link(&self) -> &Link {
        &self.link
    }

    /// Optional cached display text attached to this AI.
    pub fn text(&self) -> Option<&str> {
        self.text.as_deref()
    }

    /// Attaches cached display text, moving the binding for concise builders.
    #[must_use]
    pub fn with_text(mut self, text: impl Into<String>) -> Self {
        self.text = Some(text.into());
        self
    }

    pub(super) fn set_text(&mut self, text: String) -> Result<()> {
        if self.text.is_some() {
            return Err(Error::InvalidModel {
                field: "AI",
                reason: "one AI has more than one SeriesText",
            });
        }
        self.text = Some(text);
        Ok(())
    }
}

/// The four mandatory AI bindings of a regular series, in wire order.
///
/// Named fields make missing, duplicated, or reordered roles unrepresentable
/// after construction.
#[derive(Debug, PartialEq, Eq)]
pub struct Ai {
    name: Binding,
    values: Binding,
    categories: Binding,
    bubbles: Binding,
}

impl Ai {
    /// Creates a complete AI set and verifies each binding's semantic role.
    pub fn new(
        name: Binding,
        values: Binding,
        categories: Binding,
        bubbles: Binding,
    ) -> Result<Self> {
        for (binding, role) in [
            (&name, Role::Name),
            (&values, Role::Values),
            (&categories, Role::Categories),
            (&bubbles, Role::Bubbles),
        ] {
            if binding.link.role() != role {
                return Err(Error::InvalidModel {
                    field: "AI",
                    reason: "binding role does not match its named AI slot",
                });
            }
        }
        Ok(Self {
            name,
            values,
            categories,
            bubbles,
        })
    }

    /// Creates four automatic bindings for a producer context.
    pub fn automatic(context: Context) -> Self {
        fn link(context: Context, role: Role) -> Link {
            match context.kind() {
                Kind::Graph => Link::graph(role, Source::Automatic, RowCol::ZERO),
                Kind::Excel => Link::excel(role, Source::Automatic, Vec::new()),
            }
        }
        Self {
            name: Binding::new(link(context, Role::Name), None),
            values: Binding::new(link(context, Role::Values), None),
            categories: Binding::new(link(context, Role::Categories), None),
            bubbles: Binding::new(link(context, Role::Bubbles), None),
        }
    }

    /// Looks up one binding by semantic role.
    pub const fn get(&self, role: Role) -> &Binding {
        match role {
            Role::Name => &self.name,
            Role::Values => &self.values,
            Role::Categories => &self.categories,
            Role::Bubbles => &self.bubbles,
        }
    }

    /// Replaces one binding, selecting its named slot from the link role.
    pub fn set(&mut self, binding: Binding) -> &mut Self {
        self.replace(binding);
        self
    }

    /// Replaces one binding and returns the moved AI set for struct builders.
    #[must_use]
    pub fn with(mut self, binding: Binding) -> Self {
        self.replace(binding);
        self
    }

    pub(super) fn get_mut(&mut self, role: Role) -> &mut Binding {
        match role {
            Role::Name => &mut self.name,
            Role::Values => &mut self.values,
            Role::Categories => &mut self.categories,
            Role::Bubbles => &mut self.bubbles,
        }
    }

    pub(super) fn ordered(&self) -> [&Binding; 4] {
        [&self.name, &self.values, &self.categories, &self.bubbles]
    }

    pub(super) fn replace(&mut self, binding: Binding) {
        let role = binding.link.role();
        *self.get_mut(role) = binding;
    }
}

/// One chart series.
#[derive(Debug, PartialEq, Eq)]
pub struct Series {
    pub category_kind: DataKind,
    pub category_count: Count,
    pub value_count: Count,
    pub bubble_count: Count,
    /// Exactly one regular-group or auxiliary-series owner.
    pub owner: Owner,
    /// Exactly four AI bindings in the required semantic order.
    pub ai: Ai,
}

impl Series {
    /// Creates an empty text-category series in the primary chart group.
    pub fn new(context: Context) -> Self {
        Self {
            category_kind: DataKind::Text,
            category_count: Count::ZERO,
            value_count: Count::ZERO,
            bubble_count: Count::ZERO,
            owner: Owner::Group(GroupId::ZERO),
            ai: Ai::automatic(context),
        }
    }
}

/// Exactly one owner branch from the BIFF `SERIESFORMAT` grammar.
#[derive(Debug, PartialEq, Eq)]
pub enum Owner {
    /// Regular series assigned to one chart group by `SerToCrt`.
    Group(GroupId),
    /// Trendline assigned to a one-based parent series.
    Trend {
        parent: crate::record::series::Parent,
        /// Exact inert `SerAuxTrend` payload.
        data: [u8; 28],
    },
    /// Error bar assigned to a one-based parent series.
    ErrorBar {
        parent: crate::record::series::Parent,
        /// Exact inert `SerAuxErrBar` payload.
        data: [u8; 14],
    },
}

impl Owner {
    /// Regular primary chart-group ownership.
    pub const PRIMARY: Self = Self::Group(GroupId::ZERO);

    /// Returns the regular chart group, or `None` for an auxiliary series.
    pub const fn group(&self) -> Option<GroupId> {
        match self {
            Self::Group(group) => Some(*group),
            Self::Trend { .. } | Self::ErrorBar { .. } => None,
        }
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
    /// Axis-parent collection that owns this group.
    pub parent: axis::ParentId,
    pub order: Order,
    pub vary_colors: bool,
    pub family: Family,
    /// Excel-mandatory written-but-unused CrtLink owned by this chart group.
    ///
    /// Standalone Graph preserves this record when present but does not require
    /// it without the unavailable normative chart-sheet grammar.
    pub link: crate::record::line::Link,
    pub lines: Vec<group::Line>,
    pub drop_bars: Vec<group::DropBar>,
}

impl Group {
    /// Primary line-chart group used by a new chart.
    pub const fn line() -> Self {
        Self {
            parent: axis::ParentId::PRIMARY,
            order: Order::ZERO,
            vary_colors: false,
            family: Family::Line { flags: 0 },
            link: crate::record::line::Link::new([0; 10]),
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

/// Excel cached value, including the producer-specific `BoolErr` union.
#[derive(Debug, PartialEq)]
pub enum XlValue {
    Number(f64),
    Text(String),
    Bool(bool),
    Error(cache::Fault),
    Blank,
}

impl From<Value> for XlValue {
    fn from(value: Value) -> Self {
        match value {
            Value::Number(value) => Self::Number(value),
            Value::Text(value) => Self::Text(value),
            Value::Blank => Self::Blank,
        }
    }
}

/// Borrowed producer-neutral view of a cached value.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ValueRef<'a> {
    Number(f64),
    Text(&'a str),
    Bool(bool),
    Error(cache::Fault),
    Blank,
}

/// Producer-typed cached chart cell.
///
/// Each variant owns its producer-specific coordinate, section, and format;
/// Graph/Excel mixtures are therefore unrepresentable.
#[derive(Debug, PartialEq)]
pub enum Cache {
    /// Excel SERIESDATA cell.
    Excel {
        section: cache::Index,
        row: u16,
        col: u8,
        xf: cache::Xf,
        value: XlValue,
    },
    /// Standalone Graph datasheet cell.
    Graph {
        row: RowCol,
        col: RowCol,
        ifmt: cache::Ifmt,
        value: Value,
    },
}

impl Cache {
    /// Creates an Excel cache cell.
    pub fn excel(
        section: cache::Index,
        row: u16,
        col: u8,
        xf: cache::Xf,
        value: impl Into<XlValue>,
    ) -> Self {
        Self::Excel {
            section,
            row,
            col,
            xf,
            value: value.into(),
        }
    }

    /// Creates a standalone Graph cache cell.
    pub const fn graph(row: RowCol, col: RowCol, ifmt: cache::Ifmt, value: Value) -> Self {
        Self::Graph {
            row,
            col,
            ifmt,
            value,
        }
    }

    /// Producer grammar owned by this cell.
    pub const fn kind(&self) -> Kind {
        match self {
            Self::Excel { .. } => Kind::Excel,
            Self::Graph { .. } => Kind::Graph,
        }
    }

    /// Cached value.
    pub fn value(&self) -> ValueRef<'_> {
        match self {
            Self::Excel { value, .. } => match value {
                XlValue::Number(value) => ValueRef::Number(*value),
                XlValue::Text(value) => ValueRef::Text(value),
                XlValue::Bool(value) => ValueRef::Bool(*value),
                XlValue::Error(value) => ValueRef::Error(*value),
                XlValue::Blank => ValueRef::Blank,
            },
            Self::Graph { value, .. } => match value {
                Value::Number(value) => ValueRef::Number(*value),
                Value::Text(value) => ValueRef::Text(value),
                Value::Blank => ValueRef::Blank,
            },
        }
    }
}

/// Mutable slice guard that preserves pristine replay until mutation occurs.
pub struct Edit<'a, T> {
    values: &'a mut [T],
    dirty: &'a mut bool,
    parsed: bool,
}

impl<T> Deref for Edit<'_, T> {
    type Target = [T];

    fn deref(&self) -> &Self::Target {
        self.values
    }
}

impl<T> DerefMut for Edit<'_, T> {
    fn deref_mut(&mut self) -> &mut Self::Target {
        if self.parsed {
            *self.dirty = true;
        }
        self.values
    }
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
    pub(super) zoom: layout::Zoom,
    pub(super) growth: layout::Growth,
    pub(super) title: Option<String>,
    pub(super) series: Vec<Series>,
    pub(super) groups: Vec<Group>,
    pub(super) axes: Vec<axis::Axis>,
    pub(super) parents: Vec<axis::Parent>,
    pub(super) legend: Option<Legend>,
    pub(super) caches: Vec<Cache>,
    pub(super) dimensions: cache::Dims,
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
        let mut parents = Vec::new();
        parents
            .try_reserve_exact(1)
            .map_err(|_| Error::Allocation {
                resource: "axis parents",
            })?;
        parents.push(axis::Parent::primary(layout::Pos::default()));
        Ok(Self {
            context,
            rect: Rect::default(),
            props: Props::default(),
            zoom: layout::Zoom::default(),
            growth: layout::Growth::default(),
            title: None,
            series: Vec::new(),
            groups,
            axes: Vec::new(),
            parents,
            legend: None,
            caches: Vec::new(),
            dimensions: cache::Dims::empty(context.kind()),
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

    /// Chart-window zoom.
    pub const fn zoom(&self) -> layout::Zoom {
        self.zoom
    }

    /// Sets the checked chart-window zoom.
    pub fn set_zoom(&mut self, value: layout::Zoom) {
        self.touch();
        self.zoom = value;
    }

    /// Plot-area font growth factors.
    pub const fn growth(&self) -> layout::Growth {
        self.growth
    }

    /// Sets plot-area font growth factors.
    pub fn set_growth(&mut self, value: layout::Growth) {
        self.touch();
        self.growth = value;
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

    /// Mutably borrows series and marks parsed input dirty only on mutation.
    pub fn series_mut(&mut self) -> Edit<'_, Series> {
        Edit {
            values: &mut self.series,
            dirty: &mut self.dirty,
            parsed: matches!(&self.origin, Origin::Parsed(_)),
        }
    }

    pub fn add_series(&mut self, value: Series) -> Result<()> {
        match &value.owner {
            Owner::Group(group) if usize::from(group.get()) >= self.groups.len() => {
                return Err(Error::InvalidModel {
                    field: "series",
                    reason: "series refers to a missing chart group",
                });
            },
            Owner::Trend { parent, .. } | Owner::ErrorBar { parent, .. } => {
                let zero_based = usize::try_from(parent.series().get() - 1).map_err(|_| {
                    Error::InvalidModel {
                        field: "series",
                        reason: "auxiliary parent index exceeds this platform",
                    }
                })?;
                if self
                    .series
                    .get(zero_based)
                    .is_none_or(|parent| !matches!(parent.owner, Owner::Group(_)))
                {
                    return Err(Error::InvalidModel {
                        field: "series",
                        reason: "auxiliary series must refer to an existing regular series",
                    });
                }
            },
            _ => {},
        }
        for binding in value.ai.ordered() {
            if !matches!(
                (self.context.kind(), binding.link()),
                (Kind::Excel, Link::Excel { .. }) | (Kind::Graph, Link::Graph { .. })
            ) {
                return Err(Error::InvalidModel {
                    field: "link",
                    reason: "series binding does not match the chart producer",
                });
            }
        }
        check_add(self.series.len(), self.limits.max_series, "series count")?;
        reserve_one(&mut self.series, "chart series")?;
        self.touch();
        self.series.push(value);
        Ok(())
    }

    /// Removes an unreferenced series and retargets later auxiliary parents.
    pub fn remove_series(&mut self, index: usize) -> Result<Option<Series>> {
        if index >= self.series.len() {
            return Ok(None);
        }
        let one_based = index.checked_add(1).ok_or(Error::SizeOverflow {
            resource: "series index",
        })?;
        let one_based = u32::try_from(one_based).map_err(|_| Error::InvalidModel {
            field: "series",
            reason: "series index exceeds the auxiliary-parent range",
        })?;
        for series in &self.series {
            let parent = match &series.owner {
                Owner::Trend { parent, .. } | Owner::ErrorBar { parent, .. } => parent,
                Owner::Group(_) => continue,
            };
            let zero_based =
                usize::try_from(parent.series().get() - 1).map_err(|_| Error::InvalidModel {
                    field: "series",
                    reason: "auxiliary parent index exceeds this platform",
                })?;
            if self
                .series
                .get(zero_based)
                .is_none_or(|parent| !matches!(parent.owner, Owner::Group(_)))
            {
                return Err(Error::InvalidModel {
                    field: "series",
                    reason: "auxiliary series refers to an invalid parent",
                });
            }
        }
        if self.series.iter().any(|series| match &series.owner {
            Owner::Trend { parent, .. } | Owner::ErrorBar { parent, .. } => {
                parent.series().get() == one_based
            },
            Owner::Group(_) => false,
        }) {
            return Err(Error::InvalidModel {
                field: "series",
                reason: "series is still referenced by an auxiliary series",
            });
        }
        for series in &mut self.series {
            let parent = match &mut series.owner {
                Owner::Trend { parent, .. } | Owner::ErrorBar { parent, .. } => parent,
                Owner::Group(_) => continue,
            };
            if parent.series().get() > one_based {
                let shifted =
                    u16::try_from(parent.series().get() - 1).map_err(|_| Error::InvalidModel {
                        field: "series",
                        reason: "auxiliary parent index exceeds its checked range",
                    })?;
                *parent = crate::record::series::Parent::try_new(shifted).map_err(|_| {
                    Error::InvalidModel {
                        field: "series",
                        reason: "auxiliary parent index became invalid",
                    }
                })?;
            }
        }
        self.touch();
        Ok(Some(self.series.remove(index)))
    }

    pub fn groups(&self) -> &[Group] {
        &self.groups
    }

    /// Mutably borrows chart groups and marks parsed input dirty only on mutation.
    pub fn groups_mut(&mut self) -> Edit<'_, Group> {
        Edit {
            values: &mut self.groups,
            dirty: &mut self.dirty,
            parsed: matches!(&self.origin, Origin::Parsed(_)),
        }
    }

    pub fn add_group(&mut self, value: Group) -> Result<()> {
        check_add(self.groups.len(), self.limits.max_groups, "group count")?;
        if self
            .parents
            .get(value.parent.index())
            .is_none_or(|parent| parent.id() != value.parent)
        {
            return Err(Error::InvalidModel {
                field: "group",
                reason: "chart group refers to a missing axis parent",
            });
        }
        if self.groups.iter().any(|group| group.order == value.order) {
            return Err(Error::InvalidModel {
                field: "group",
                reason: "chart-group drawing order is duplicated",
            });
        }
        reserve_one(&mut self.groups, "chart groups")?;
        self.touch();
        self.groups.push(value);
        Ok(())
    }

    /// Removes an unreferenced group and retargets later group indices.
    ///
    /// A referenced group is refused instead of silently moving its series to
    /// a different chart family.
    pub fn remove_group(&mut self, index: usize) -> Result<Option<Group>> {
        if index >= self.groups.len() {
            return Ok(None);
        }
        let raw = u8::try_from(index).map_err(|_| Error::InvalidModel {
            field: "group",
            reason: "chart-group index exceeds nine",
        })?;
        let id = GroupId::new(raw).ok_or(Error::InvalidModel {
            field: "group",
            reason: "chart-group index exceeds nine",
        })?;
        if self.series.iter().any(|series| {
            series
                .owner
                .group()
                .is_some_and(|group| usize::from(group.get()) >= self.groups.len())
        }) {
            return Err(Error::InvalidModel {
                field: "series",
                reason: "series refers to an invalid chart group",
            });
        }
        if self
            .series
            .iter()
            .any(|series| series.owner.group() == Some(id))
        {
            return Err(Error::InvalidModel {
                field: "group",
                reason: "chart group is still referenced by a series",
            });
        }
        for series in &mut self.series {
            if let Owner::Group(group) = &mut series.owner
                && group.get() > raw
            {
                *group = GroupId::new(group.get() - 1).ok_or(Error::InvalidModel {
                    field: "series",
                    reason: "series chart-group index became invalid",
                })?;
            }
        }
        self.touch();
        Ok(Some(self.groups.remove(index)))
    }

    pub fn axes(&self) -> &[axis::Axis] {
        &self.axes
    }

    /// Mutably borrows axes and marks parsed input dirty only on mutation.
    pub fn axes_mut(&mut self) -> Edit<'_, axis::Axis> {
        Edit {
            values: &mut self.axes,
            dirty: &mut self.dirty,
            parsed: matches!(&self.origin, Origin::Parsed(_)),
        }
    }

    pub fn add_axis(&mut self, value: axis::Axis) -> Result<()> {
        if self
            .parents
            .get(value.parent.index())
            .is_none_or(|parent| parent.id() != value.parent)
        {
            return Err(Error::InvalidModel {
                field: "axis",
                reason: "axis refers to a missing axis parent",
            });
        }
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

    /// Primary and optional secondary axis-parent collections.
    pub fn parents(&self) -> &[axis::Parent] {
        &self.parents
    }

    /// Mutably borrows axis-parent metadata.
    pub fn parents_mut(&mut self) -> Edit<'_, axis::Parent> {
        Edit {
            values: &mut self.parents,
            dirty: &mut self.dirty,
            parsed: matches!(&self.origin, Origin::Parsed(_)),
        }
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

    /// Context-specific mandatory cache dimensions.
    pub const fn dimensions(&self) -> cache::Dims {
        self.dimensions
    }

    /// Sets producer-typed cache dimensions.
    pub fn set_dimensions(&mut self, value: cache::Dims) -> Result<()> {
        if !value.matches(self.context.kind()) {
            return Err(Error::InvalidModel {
                field: "Dimensions",
                reason: "dimensions do not match the chart producer",
            });
        }
        let derived = cache_dimensions(&self.caches, self.context.kind())?;
        if !dimensions_cover(value, derived) {
            return Err(Error::InvalidModel {
                field: "Dimensions",
                reason: "dimensions do not cover the cached chart cells",
            });
        }
        self.touch();
        self.dimensions = value;
        Ok(())
    }

    /// Mutably borrows cached cells and marks parsed input dirty only on mutation.
    /// Call [`Self::sync_dimensions`] after changing cell coordinates.
    pub fn caches_mut(&mut self) -> Edit<'_, Cache> {
        Edit {
            values: &mut self.caches,
            dirty: &mut self.dirty,
            parsed: matches!(&self.origin, Origin::Parsed(_)),
        }
    }

    pub fn add_cache(&mut self, value: Cache) -> Result<()> {
        if value.kind() != self.context.kind() {
            return Err(Error::InvalidModel {
                field: "cache",
                reason: "cached cell does not match the chart producer",
            });
        }
        check_add(
            self.caches.len(),
            self.limits.max_cached_values,
            "cached value count",
        )?;
        reserve_one(&mut self.caches, "chart cache")?;
        self.caches.push(value);
        let dimensions = match cache_dimensions(&self.caches, self.context.kind()) {
            Ok(value) => value,
            Err(error) => {
                let _ = self.caches.pop();
                return Err(error);
            },
        };
        self.touch();
        self.dimensions = dimensions;
        Ok(())
    }

    /// Removes one cached cell and synchronizes producer dimensions.
    pub fn remove_cache(&mut self, index: usize) -> Result<Option<Cache>> {
        if index >= self.caches.len() {
            return Ok(None);
        }
        let removed = self.caches.remove(index);
        let dimensions = match cache_dimensions(&self.caches, self.context.kind()) {
            Ok(value) => value,
            Err(error) => {
                self.caches.insert(index, removed);
                return Err(error);
            },
        };
        self.touch();
        self.dimensions = dimensions;
        Ok(Some(removed))
    }

    /// Recomputes mandatory dimensions from the current cached cells.
    pub fn sync_dimensions(&mut self) -> Result<()> {
        let dimensions = cache_dimensions(&self.caches, self.context.kind())?;
        self.touch();
        self.dimensions = dimensions;
        Ok(())
    }

    pub fn formats(&self) -> &[format::Format] {
        &self.formats
    }

    pub fn formats_mut(&mut self) -> Edit<'_, format::Format> {
        Edit {
            values: &mut self.formats,
            dirty: &mut self.dirty,
            parsed: matches!(&self.origin, Origin::Parsed(_)),
        }
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

    pub fn labels_mut(&mut self) -> Edit<'_, Label> {
        Edit {
            values: &mut self.labels,
            dirty: &mut self.dirty,
            parsed: matches!(&self.origin, Origin::Parsed(_)),
        }
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

pub(super) fn cache_dimensions(values: &[Cache], kind: Kind) -> Result<cache::Dims> {
    match kind {
        Kind::Excel => {
            let mut bounds: Option<(u16, u16, u8, u8)> = None;
            for value in values {
                let Cache::Excel { row, col, .. } = value else {
                    return Err(Error::InvalidModel {
                        field: "cache",
                        reason: "Graph cache cell appears in an Excel chart",
                    });
                };
                bounds = Some(match bounds {
                    Some((first_row, last_row, first_col, last_col)) => (
                        first_row.min(*row),
                        last_row.max(*row),
                        first_col.min(*col),
                        last_col.max(*col),
                    ),
                    None => (*row, *row, *col, *col),
                });
            }
            let dimensions = match bounds {
                Some((first_row, last_row, first_col, last_col)) => cache::ExcelDims::new(
                    u32::from(first_row),
                    u32::from(last_row) + 1,
                    u16::from(first_col),
                    u16::from(last_col) + 1,
                )
                .ok_or(Error::InvalidModel {
                    field: "Dimensions",
                    reason: "Excel cache bounds are outside the BIFF8 grid",
                })?,
                None => cache::ExcelDims::default(),
            };
            Ok(cache::Dims::Excel(dimensions))
        },
        Kind::Graph => {
            let mut coordinates = Vec::new();
            coordinates
                .try_reserve_exact(values.len())
                .map_err(|_| Error::Allocation {
                    resource: "Graph cache coordinates",
                })?;
            for value in values {
                let Cache::Graph { row, col, .. } = value else {
                    return Err(Error::InvalidModel {
                        field: "cache",
                        reason: "Excel cache cell appears in a Graph chart",
                    });
                };
                coordinates.push((u32::from(row.get()) << 12) | u32::from(col.get()));
            }
            coordinates.sort_unstable();
            coordinates.dedup();

            let mut current_row = None;
            let mut width = 0u16;
            let mut longest = 0u16;
            let mut rows = 0u16;
            for coordinate in coordinates {
                let row = coordinate >> 12;
                if current_row != Some(row) {
                    longest = longest.max(width);
                    width = 0;
                    rows = rows.checked_add(1).ok_or(Error::SizeOverflow {
                        resource: "Graph cache rows",
                    })?;
                    current_row = Some(row);
                }
                width = width.checked_add(1).ok_or(Error::SizeOverflow {
                    resource: "Graph cache row width",
                })?;
            }
            longest = longest.max(width);
            let rows = u8::try_from(rows).map_err(|_| Error::InvalidModel {
                field: "Dimensions",
                reason: "Graph cache has more than 255 non-empty rows",
            })?;
            let longest = RowCol::new(longest).ok_or(Error::InvalidModel {
                field: "Dimensions",
                reason: "Graph cache row has more than 3,999 cells",
            })?;
            let dimensions = cache::GraphDims::new(longest, rows).ok_or(Error::InvalidModel {
                field: "Dimensions",
                reason: "Graph cache dimensions are inconsistent",
            })?;
            Ok(cache::Dims::Graph(dimensions))
        },
    }
}

pub(super) const fn dimensions_cover(declared: cache::Dims, derived: cache::Dims) -> bool {
    match (declared, derived) {
        (cache::Dims::Excel(declared), cache::Dims::Excel(derived)) => {
            if derived.row_after() == 0 {
                declared.row_after() == 0
            } else {
                declared.row_after() != 0
                    && declared.first_row() <= derived.first_row()
                    && declared.row_after() >= derived.row_after()
                    && declared.first_col() <= derived.first_col()
                    && declared.col_after() >= derived.col_after()
            }
        },
        (cache::Dims::Graph(declared), cache::Dims::Graph(derived)) => {
            declared.longest_row().get() == derived.longest_row().get()
                && declared.rows() == derived.rows()
        },
        _ => false,
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

    fn omit(stream: &chart::Stream, target: RecordKind) -> Vec<u8> {
        let mut out = Encoder::new();
        for item in Records::new(stream.as_bytes()) {
            let record = item.expect("valid fixture record");
            if record.kind() != target {
                out.push_ref(record).expect("record replay");
            }
        }
        out.finish()
    }

    fn excel_input(bytes: &[u8]) -> chart::Ref<'_> {
        chart::Ref::open(bytes).expect("well-framed chart rewrite")
    }

    fn excel_chart() -> Chart {
        let context = Context::excel().with_external_sheets(1);
        let mut chart = Chart::new(context).expect("new chart");
        let mut series = Series::new(context);
        series.category_kind = DataKind::Text;
        series.category_count = count(2);
        series.value_count = count(2);
        series.ai = Ai::new(
            Binding::new(
                Link::excel(Role::Name, Source::Automatic, Vec::new()),
                Some("FY26".into()),
            ),
            Binding::new(
                Link::excel(
                    Role::Values,
                    Source::Cells,
                    vec![0x1B, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0],
                ),
                None,
            ),
            Binding::new(
                Link::excel(Role::Categories, Source::Automatic, Vec::new()),
                None,
            ),
            Binding::new(
                Link::excel(Role::Bubbles, Source::Automatic, Vec::new()),
                None,
            ),
        )
        .expect("canonical AI roles");
        chart.add_series(series).expect("series");
        chart
            .add_cache(Cache::excel(
                cache::Index::Values,
                0,
                0,
                cache::Xf::new(0),
                Value::Number(42.5),
            ))
            .expect("numeric cache");
        chart
            .add_cache(Cache::excel(
                cache::Index::Values,
                1,
                0,
                cache::Xf::new(2),
                Value::Text("safe".into()),
            ))
            .expect("text cache");
        chart
            .add_cache(Cache::excel(
                cache::Index::Values,
                2,
                0,
                cache::Xf::new(3),
                Value::Blank,
            ))
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
        assert_eq!(parsed.title(), None);
        assert_eq!(parsed.series().len(), 1);
        assert_eq!(parsed.caches().len(), 3);
        assert!(matches!(parsed.caches()[2].value(), ValueRef::Blank));
        let replay = parsed.encode().expect("exact replay");
        assert_eq!(replay.as_bytes().as_ptr(), pointer);
    }

    #[test]
    fn graph_and_excel_links_have_distinct_checked_wire_grammars() {
        let row_col = RowCol::new(7).expect("Graph coordinate");
        let mut graph = Chart::new(Context::graph()).expect("Graph chart");
        let mut series = Series::new(Context::graph());
        series.ai.replace(Binding::new(
            Link::graph(Role::Values, Source::Literal, row_col),
            None,
        ));
        graph.add_series(series).expect("Graph series");
        graph
            .add_cache(Cache::graph(
                RowCol::new(1).expect("row"),
                RowCol::new(2).expect("column"),
                cache::Ifmt::new(4),
                Value::Blank,
            ))
            .expect("Graph blank");
        let stream = fixture(graph);
        let parsed = Chart::open(stream, Context::graph()).expect("Graph parse");
        assert!(matches!(
            parsed.series()[0].ai.get(Role::Values).link(),
            Link::Graph { .. }
        ));
        assert!(matches!(parsed.caches()[0], Cache::Graph { .. }));

        let mut wrong = Chart::new(Context::graph()).expect("Graph chart");
        let mut wrong_series = Series::new(Context::graph());
        wrong_series.ai.replace(Binding::new(
            Link::excel(Role::Values, Source::Automatic, Vec::new()),
            None,
        ));
        assert!(matches!(
            wrong.add_series(wrong_series),
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

        let mut defined = Chart::new(Context::excel()).expect("chart");
        defined.set_props(Props {
            flags: 1 << 2,
            plot_area: true,
        });
        defined.authoring_proven = true;
        assert!(defined.encode().is_ok());

        let mut reserved = Chart::new(Context::excel()).expect("chart");
        reserved.set_props(Props {
            flags: 1 << 5,
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
        chart
            .add_series(Series::new(Context::excel()))
            .expect("first series");
        assert!(matches!(
            chart.add_series(Series::new(Context::excel())),
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
        let mut groups = chart.groups_mut();
        let group = groups.first_mut().expect("default line group");
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

    #[test]
    fn regular_series_requires_one_owner_and_series_text_remains_ai_local() {
        let stream = fixture(excel_chart());
        let missing = omit(&stream, RecordKind::new(0x1045));
        assert!(matches!(
            Chart::parse(
                excel_input(&missing),
                Context::excel().with_external_sheets(1)
            ),
            Err(Error::InvalidChart {
                reason: "Series requires exactly four AI bindings and one SerToCrt",
                ..
            })
        ));

        let mut out = Encoder::new();
        let mut ai = 0usize;
        for item in Records::new(stream.as_bytes()) {
            let record = item.expect("valid fixture record");
            out.push_ref(record).expect("record replay");
            if record.kind() == RecordKind::new(0x1051) {
                ai += 1;
                if ai == Role::ALL.len() {
                    let text = [0, 0, 1, 0, b'x'];
                    out.push(RecordKind::new(0x100D), &text)
                        .expect("first optional SeriesText");
                    out.push(RecordKind::new(0x100D), &text)
                        .expect("misplaced second SeriesText");
                }
            }
        }
        let duplicate = out.finish();
        assert!(matches!(
            Chart::parse(
                excel_input(&duplicate),
                Context::excel().with_external_sheets(1)
            ),
            Err(Error::InvalidChart {
                reason: "SeriesText in Series must immediately follow one BRAI",
                ..
            })
        ));
    }

    #[test]
    fn auxiliary_owner_round_trips_and_series_removal_preserves_dependencies() {
        let context = Context::excel().with_external_sheets(1);
        let mut chart = excel_chart();
        let mut auxiliary = Series::new(context);
        auxiliary.owner = Owner::Trend {
            parent: crate::record::series::Parent::try_new(1).expect("parent series"),
            data: [0; 28],
        };
        chart.add_series(auxiliary).expect("auxiliary series");
        let parsed = Chart::open(fixture(chart), context).expect("auxiliary parse");
        assert!(matches!(parsed.series()[1].owner, Owner::Trend { .. }));

        let mut blocked = excel_chart();
        let mut auxiliary = Series::new(context);
        auxiliary.owner = Owner::ErrorBar {
            parent: crate::record::series::Parent::try_new(1).expect("parent series"),
            data: [0; 14],
        };
        blocked.add_series(auxiliary).expect("auxiliary series");
        assert!(matches!(
            blocked.remove_series(0),
            Err(Error::InvalidModel {
                field: "series",
                ..
            })
        ));
        assert_eq!(blocked.series().len(), 2);

        let mut shifted = Chart::new(context).expect("chart");
        shifted
            .add_series(Series::new(context))
            .expect("first regular series");
        shifted
            .add_series(Series::new(context))
            .expect("second regular series");
        let mut auxiliary = Series::new(context);
        auxiliary.owner = Owner::Trend {
            parent: crate::record::series::Parent::try_new(2).expect("second parent"),
            data: [0; 28],
        };
        shifted.add_series(auxiliary).expect("auxiliary series");
        assert!(shifted.remove_series(0).expect("safe removal").is_some());
        let Owner::Trend { parent, .. } = &shifted.series()[1].owner else {
            panic!("shifted auxiliary owner");
        };
        assert_eq!(parent.series().get(), 1);
    }

    #[test]
    fn cache_dimensions_and_bool_err_follow_the_typed_crud_model() {
        let mut chart = excel_chart();
        chart
            .add_cache(Cache::excel(
                cache::Index::Values,
                3,
                0,
                cache::Xf::new(4),
                XlValue::Bool(true),
            ))
            .expect("Boolean cache");
        chart
            .add_cache(Cache::excel(
                cache::Index::Values,
                4,
                0,
                cache::Xf::new(5),
                XlValue::Error(cache::Fault::DivZero),
            ))
            .expect("error cache");
        assert!(matches!(
            chart.dimensions(),
            cache::Dims::Excel(value) if value.row_after() == 5 && value.col_after() == 1
        ));
        assert!(
            chart
                .set_dimensions(cache::Dims::Excel(
                    cache::ExcelDims::new(0, 4, 0, 1).expect("smaller range")
                ))
                .is_err()
        );

        let parsed = Chart::open(fixture(chart), Context::excel().with_external_sheets(1))
            .expect("BoolErr round trip");
        assert_eq!(parsed.caches()[3].value(), ValueRef::Bool(true));
        assert_eq!(
            parsed.caches()[4].value(),
            ValueRef::Error(cache::Fault::DivZero)
        );
    }

    #[test]
    fn excel_cache_label_uses_xl_unicode_string_wire_format() {
        let text = "界".repeat(300);
        let mut chart = excel_chart();
        {
            let mut caches = chart.caches_mut();
            let Cache::Excel { value, .. } = caches.get_mut(1).expect("text cache") else {
                panic!("expected Excel text cache");
            };
            *value = XlValue::Text(text.clone());
        }

        let stream = fixture(chart);
        let label = stream
            .records()
            .map(|record| record.expect("valid fixture record"))
            .find(|record| record.kind() == RecordKind::new(0x0204))
            .expect("Label record");
        assert_eq!(
            label.payload().get(6..9),
            Some([0x2C, 0x01, 0x01].as_slice())
        );
        assert_eq!(label.payload().len(), 6 + 3 + 300 * 2);

        let parsed = Chart::open(stream, Context::excel().with_external_sheets(1))
            .expect("XLUnicodeString Label round trip");
        assert_eq!(parsed.caches()[1].value(), ValueRef::Text(&text));
    }

    #[test]
    fn group_and_parent_crud_preserves_semantic_ownership() {
        let context = Context::excel();
        let mut chart = Chart::new(context).expect("chart");
        let mut second = Group::line();
        second.order = Order::new(1).expect("drawing order");
        chart.add_group(second).expect("second group");
        let mut series = Series::new(context);
        series.owner = Owner::Group(GroupId::new(1).expect("second group index"));
        chart.add_series(series).expect("series");

        assert!(matches!(
            chart.remove_group(1),
            Err(Error::InvalidModel { field: "group", .. })
        ));
        assert!(chart.remove_group(0).expect("safe removal").is_some());
        assert_eq!(chart.groups().len(), 1);
        assert_eq!(chart.series()[0].owner.group(), Some(GroupId::ZERO));

        chart
            .add_axis(axis::Axis::new(axis::Kind::Category))
            .expect("primary axis");
        let parsed = Chart::open(fixture(chart), context).expect("parent ownership parse");
        assert_eq!(parsed.groups()[0].parent, axis::ParentId::PRIMARY);
        assert_eq!(parsed.axes()[0].parent, axis::ParentId::PRIMARY);
    }

    #[test]
    fn mutable_borrow_marks_parsed_input_dirty_only_after_write() {
        let context = Context::excel().with_external_sheets(1);
        let mut chart = Chart::open(fixture(excel_chart()), context).expect("parsed chart");
        {
            let groups = chart.groups_mut();
            assert_eq!(groups.len(), 1);
        }
        assert!(chart.is_pristine());
        {
            let mut groups = chart.groups_mut();
            groups[0].vary_colors = true;
        }
        assert!(!chart.is_pristine());
    }

    #[test]
    fn excel_rejects_proven_topology_violations_and_bad_siindex_order() {
        let context = Context::excel().with_external_sheets(1);
        let stream = fixture(excel_chart());
        for kind in [0x00A0, 0x1022, 0x104F].map(RecordKind::new) {
            let malformed = omit(&stream, kind);
            assert!(Chart::parse(excel_input(&malformed), context).is_err());
        }

        let mut out = Encoder::new();
        let mut section = 0usize;
        for item in Records::new(stream.as_bytes()) {
            let record = item.expect("valid fixture record");
            if record.kind() == RecordKind::new(0x1065) {
                section += 1;
                if section == 2 {
                    out.push(record.kind(), &3u16.to_le_bytes())
                        .expect("out-of-order SIIndex");
                    continue;
                }
            }
            out.push_ref(record).expect("record replay");
        }
        let malformed = out.finish();
        assert!(matches!(
            Chart::parse(excel_input(&malformed), context),
            Err(Error::InvalidChart {
                reason: "SIIndex sections are missing, duplicated, or out of order",
                ..
            })
        ));

        let mut out = Encoder::new();
        for item in Records::new(stream.as_bytes()) {
            let record = item.expect("valid fixture record");
            if record.kind() == RecordKind::new(0x0200) {
                out.push(RecordKind::new(0x1033), &[])
                    .expect("orphan Begin");
                out.push(RecordKind::new(0x1034), &[])
                    .expect("balanced orphan End");
            }
            out.push_ref(record).expect("record replay");
        }
        let malformed = out.finish();
        assert!(matches!(
            Chart::parse(excel_input(&malformed), context),
            Err(Error::InvalidChart {
                reason: "Begin record has no chart-level collection owner",
                ..
            })
        ));
    }

    #[test]
    fn graph_does_not_inherit_excel_outer_order_but_rejects_siindex() {
        let stream = fixture(Chart::new(Context::graph()).expect("Graph chart"));
        let without_excel_scl = omit(&stream, RecordKind::new(0x00A0));
        let parsed = Chart::parse(
            chart::Ref::open(&without_excel_scl).expect("Graph rewrite"),
            Context::graph(),
        )
        .expect("Graph does not require Excel CHARTFOMATS order");
        assert!(parsed.is_pristine());
        assert_eq!(
            parsed.encode().expect("exact Graph replay").as_bytes(),
            without_excel_scl
        );

        let without_excel_crt_link = omit(&stream, RecordKind::new(0x1022));
        let parsed = Chart::parse(
            chart::Ref::open(&without_excel_crt_link).expect("Graph rewrite"),
            Context::graph(),
        )
        .expect("Graph does not require the Excel-mandatory CrtLink");
        assert_eq!(
            parsed.encode().expect("exact Graph replay").as_bytes(),
            without_excel_crt_link
        );

        let mut out = Encoder::new();
        for item in Records::new(stream.as_bytes()) {
            let record = item.expect("valid fixture record");
            if record.kind() == super::super::EOF {
                out.push(RecordKind::new(0x1065), &1u16.to_le_bytes())
                    .expect("Graph SIIndex");
            }
            out.push_ref(record).expect("record replay");
        }
        let malformed = out.finish();
        let input = chart::Ref::open(&malformed).expect("Graph SIIndex framing");
        assert!(matches!(
            Chart::parse(input, Context::graph()),
            Err(Error::InvalidChart {
                reason: "SIIndex is not part of the standalone Graph grammar",
                ..
            })
        ));
    }
}
