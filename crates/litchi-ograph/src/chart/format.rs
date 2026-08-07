//! Chart line, area, marker, data-point, and pie formatting.

/// Line appearance, including producer flags and color index.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Line {
    /// RGB or palette-derived color bytes as stored by BIFF.
    pub color: [u8; 4],
    /// Line pattern identifier.
    pub pattern: u16,
    /// Line weight identifier.
    pub weight: i16,
    /// Raw formatting flags.
    pub flags: u16,
    /// Indexed color fallback.
    pub color_index: u16,
}

/// Filled-area appearance, including palette fallbacks.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Area {
    /// Foreground color bytes.
    pub foreground: [u8; 4],
    /// Background color bytes.
    pub background: [u8; 4],
    /// Fill pattern identifier.
    pub pattern: u16,
    /// Raw formatting flags.
    pub flags: u16,
    /// Indexed foreground fallback.
    pub foreground_index: u16,
    /// Indexed background fallback.
    pub background_index: u16,
}

/// One formatting record retained in chart order.
#[derive(Debug, PartialEq, Eq)]
pub enum Format {
    /// Line formatting.
    Line(Line),
    /// Area formatting.
    Area(Area),
    /// Opaque `MarkerFormat` payload.
    Marker {
        /// Exact `MarkerFormat` record payload.
        data: Vec<u8>,
    },
    /// Data-point or whole-series formatting selector.
    Data {
        /// Data-point index.
        point: u16,
        /// Series index.
        series: u16,
        /// Raw `DataFormat` flags.
        flags: u16,
    },
    /// Pie data-point explosion distance.
    Pie {
        /// Stored explosion value.
        explosion: u16,
    },
}
