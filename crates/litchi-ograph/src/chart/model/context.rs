use super::super::Kind;

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
