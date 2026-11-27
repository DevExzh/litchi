/// A nonnegative drawing-grid spacing in twips within the RTF signed-16-bit range.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DrawingGridSpacing(u16);

impl DrawingGridSpacing {
    /// Construct a spacing from `0` through `32767` twips.
    #[must_use]
    pub const fn new(value: u16) -> Option<Self> {
        if value <= i16::MAX as u16 {
            Some(Self(value))
        } else {
            None
        }
    }

    /// Return the spacing in twips.
    #[must_use]
    pub const fn get(self) -> u16 {
        self.0
    }
}

/// A nonnegative interval for displaying every Nth drawing-grid line.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DrawingGridLineInterval(u16);

impl DrawingGridLineInterval {
    /// Construct an interval from `0` through `32767`.
    #[must_use]
    pub const fn new(value: u16) -> Option<Self> {
        if value <= i16::MAX as u16 {
            Some(Self(value))
        } else {
            None
        }
    }

    /// Return the display interval.
    #[must_use]
    pub const fn get(self) -> u16 {
        self.0
    }
}

/// Passive document-level drawing-grid controls.
///
/// These values are retained for round-tripping only. This crate does not
/// render a grid, align content to it, or perform snapping.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DocumentDrawingGrid {
    /// Whether the grid follows document margins (`dgmargin`).
    pub follows_margins: bool,
    /// Whether snapping to the drawing grid was requested (`dgsnap`).
    pub snap_to_grid: bool,
    pub horizontal_spacing: Option<DrawingGridSpacing>,
    pub vertical_spacing: Option<DrawingGridSpacing>,
    /// Horizontal grid origin in twips, within the RTF signed-16-bit range.
    pub horizontal_origin_twips: Option<i16>,
    /// Vertical grid origin in twips, within the RTF signed-16-bit range.
    pub vertical_origin_twips: Option<i16>,
    pub horizontal_line_interval: Option<DrawingGridLineInterval>,
    pub vertical_line_interval: Option<DrawingGridLineInterval>,
}

impl DocumentDrawingGrid {
    pub const DEFAULT_HORIZONTAL_SPACING: DrawingGridSpacing = DrawingGridSpacing(120);
    pub const DEFAULT_VERTICAL_SPACING: DrawingGridSpacing = DrawingGridSpacing(120);
    pub const DEFAULT_HORIZONTAL_ORIGIN_TWIPS: i16 = 1701;
    pub const DEFAULT_VERTICAL_ORIGIN_TWIPS: i16 = 1984;
    pub const DEFAULT_HORIZONTAL_LINE_INTERVAL: DrawingGridLineInterval =
        DrawingGridLineInterval(3);
    pub const DEFAULT_VERTICAL_LINE_INTERVAL: DrawingGridLineInterval = DrawingGridLineInterval(0);

    /// Return whether every drawing-grid control was omitted.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        !self.follows_margins
            && !self.snap_to_grid
            && self.horizontal_spacing.is_none()
            && self.vertical_spacing.is_none()
            && self.horizontal_origin_twips.is_none()
            && self.vertical_origin_twips.is_none()
            && self.horizontal_line_interval.is_none()
            && self.vertical_line_interval.is_none()
    }

    #[must_use]
    pub fn effective_horizontal_spacing(&self) -> DrawingGridSpacing {
        self.horizontal_spacing
            .unwrap_or(Self::DEFAULT_HORIZONTAL_SPACING)
    }

    #[must_use]
    pub fn effective_vertical_spacing(&self) -> DrawingGridSpacing {
        self.vertical_spacing
            .unwrap_or(Self::DEFAULT_VERTICAL_SPACING)
    }

    #[must_use]
    pub fn effective_horizontal_origin_twips(&self) -> i16 {
        self.horizontal_origin_twips
            .unwrap_or(Self::DEFAULT_HORIZONTAL_ORIGIN_TWIPS)
    }

    #[must_use]
    pub fn effective_vertical_origin_twips(&self) -> i16 {
        self.vertical_origin_twips
            .unwrap_or(Self::DEFAULT_VERTICAL_ORIGIN_TWIPS)
    }

    #[must_use]
    pub fn effective_horizontal_line_interval(&self) -> DrawingGridLineInterval {
        self.horizontal_line_interval
            .unwrap_or(Self::DEFAULT_HORIZONTAL_LINE_INTERVAL)
    }

    #[must_use]
    pub fn effective_vertical_line_interval(&self) -> DrawingGridLineInterval {
        self.vertical_line_interval
            .unwrap_or(Self::DEFAULT_VERTICAL_LINE_INTERVAL)
    }
}
