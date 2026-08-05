use super::super::Kind;
use super::context::{Context, Count, GroupId};
use crate::{Error, Result};

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

    pub(in crate::chart) fn set_text(&mut self, text: String) -> Result<()> {
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

    pub(in crate::chart) fn get_mut(&mut self, role: Role) -> &mut Binding {
        match role {
            Role::Name => &mut self.name,
            Role::Values => &mut self.values,
            Role::Categories => &mut self.categories,
            Role::Bubbles => &mut self.bubbles,
        }
    }

    pub(in crate::chart) fn ordered(&self) -> [&Binding; 4] {
        [&self.name, &self.values, &self.categories, &self.bubbles]
    }

    pub(in crate::chart) fn replace(&mut self, binding: Binding) {
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
