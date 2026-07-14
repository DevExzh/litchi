//! Main chart structure.
//!
//! This module contains the top-level chart structure that combines
//! all chart elements (plot area, legend, title, etc.).

use crate::charts::legend::Legend;
use crate::charts::models::{Layout, TitleText};
use crate::charts::plot_area::PlotArea;
use crate::charts::series::{DataLabel, Marker};
use crate::charts::types::DisplayBlanks;

/// Printed chart header and footer strings and selection flags.
#[derive(Debug, Clone)]
pub struct ChartHeaderFooter {
    /// Odd-page header
    pub odd_header: Option<String>,
    /// Odd-page footer
    pub odd_footer: Option<String>,
    /// Even-page header
    pub even_header: Option<String>,
    /// Even-page footer
    pub even_footer: Option<String>,
    /// First-page header
    pub first_header: Option<String>,
    /// First-page footer
    pub first_footer: Option<String>,
    /// Align header and footer with page margins
    pub align_with_margins: bool,
    /// Use distinct odd- and even-page strings
    pub different_odd_even: bool,
    /// Use distinct first-page strings
    pub different_first: bool,
}

impl ChartHeaderFooter {
    /// Create an empty header/footer using schema defaults.
    #[inline]
    pub fn new() -> Self {
        Self {
            odd_header: None,
            odd_footer: None,
            even_header: None,
            even_footer: None,
            first_header: None,
            first_footer: None,
            align_with_margins: true,
            different_odd_even: false,
            different_first: false,
        }
    }
}

impl Default for ChartHeaderFooter {
    #[inline]
    fn default() -> Self {
        Self::new()
    }
}

/// Page margins for printing a chart, in inches.
#[derive(Debug, Clone, Copy)]
pub struct ChartPageMargins {
    /// Left margin
    pub left: f64,
    /// Right margin
    pub right: f64,
    /// Top margin
    pub top: f64,
    /// Bottom margin
    pub bottom: f64,
    /// Header margin
    pub header: f64,
    /// Footer margin
    pub footer: f64,
}

impl ChartPageMargins {
    /// Create a complete page-margin set.
    #[inline]
    pub fn new(left: f64, right: f64, top: f64, bottom: f64, header: f64, footer: f64) -> Self {
        Self {
            left,
            right,
            top,
            bottom,
            header,
            footer,
        }
    }
}

/// Printed chart page orientation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ChartPageOrientation {
    /// Use the printer default
    #[default]
    Default,
    /// Portrait orientation
    Portrait,
    /// Landscape orientation
    Landscape,
}

impl ChartPageOrientation {
    pub(crate) const fn xml_value(self) -> &'static str {
        match self {
            Self::Default => "default",
            Self::Portrait => "portrait",
            Self::Landscape => "landscape",
        }
    }
}

/// Page setup used when printing a chart.
#[derive(Debug, Clone, Copy)]
pub struct ChartPageSetup {
    /// Printer paper-size code
    pub paper_size: u32,
    /// First printed page number
    pub first_page_number: u32,
    /// Page orientation
    pub orientation: ChartPageOrientation,
    /// Print in black and white
    pub black_and_white: bool,
    /// Use draft quality
    pub draft: bool,
    /// Honor `first_page_number`
    pub use_first_page_number: bool,
    /// Horizontal printer resolution
    pub horizontal_dpi: i32,
    /// Vertical printer resolution
    pub vertical_dpi: i32,
    /// Number of copies
    pub copies: u32,
}

impl ChartPageSetup {
    /// Create page setup using schema defaults.
    #[inline]
    pub fn new() -> Self {
        Self {
            paper_size: 1,
            first_page_number: 1,
            orientation: ChartPageOrientation::Default,
            black_and_white: false,
            draft: false,
            use_first_page_number: false,
            horizontal_dpi: 600,
            vertical_dpi: 600,
            copies: 1,
        }
    }
}

impl Default for ChartPageSetup {
    #[inline]
    fn default() -> Self {
        Self::new()
    }
}

/// Optional chart printing configuration.
#[derive(Debug, Clone, Default)]
pub struct ChartPrintSettings {
    /// Header and footer settings
    pub header_footer: Option<ChartHeaderFooter>,
    /// Page margins
    pub page_margins: Option<ChartPageMargins>,
    /// Page setup
    pub page_setup: Option<ChartPageSetup>,
}

impl ChartPrintSettings {
    /// Create an explicitly empty print-settings container.
    #[inline]
    pub fn new() -> Self {
        Self::default()
    }
}

/// Formatting override for one pivot-chart data point.
#[derive(Debug, Clone)]
pub struct PivotFormat {
    /// Zero-based data-point index
    pub index: u32,
    /// Optional marker override
    pub marker: Option<Marker>,
    /// Optional data-label override
    pub data_label: Option<DataLabel>,
}

/// Pivot-table source metadata for a pivot chart.
#[derive(Debug, Clone)]
pub struct PivotSource {
    /// Pivot-table name
    pub name: String,
    /// Pivot format identifier
    pub format_id: u32,
}

impl PivotSource {
    /// Create pivot-source metadata.
    #[inline]
    pub fn new(name: impl Into<String>, format_id: u32) -> Self {
        Self {
            name: name.into(),
            format_id,
        }
    }
}

impl PivotFormat {
    /// Create a pivot-format entry for one data point.
    #[inline]
    pub fn new(index: u32) -> Self {
        Self {
            index,
            marker: None,
            data_label: None,
        }
    }
}

/// View 3D settings for 3D charts.
#[derive(Debug, Clone)]
pub struct View3D {
    /// Rotation around X axis (0-360 degrees)
    pub rot_x: Option<u32>,
    /// Rotation around Y axis (0-360 degrees)
    pub rot_y: Option<u32>,
    /// Right-angle axes
    pub right_angle_axes: bool,
    /// Perspective (0-240)
    pub perspective: Option<u32>,
    /// Height percent (5-500%)
    pub height_percent: Option<u32>,
    /// Depth percent (20-2000%)
    pub depth_percent: Option<u32>,
}

impl View3D {
    /// Create a new 3D view with default settings.
    #[inline]
    pub fn new() -> Self {
        Self {
            rot_x: None,
            rot_y: None,
            right_angle_axes: true,
            perspective: None,
            height_percent: None,
            depth_percent: None,
        }
    }

    /// Set rotation angles.
    #[inline]
    pub fn with_rotation(mut self, rot_x: u32, rot_y: u32) -> Self {
        self.rot_x = Some(rot_x);
        self.rot_y = Some(rot_y);
        self
    }

    /// Set perspective.
    #[inline]
    pub fn with_perspective(mut self, perspective: u32) -> Self {
        self.perspective = Some(perspective);
        self
    }
}

impl Default for View3D {
    #[inline]
    fn default() -> Self {
        Self::new()
    }
}

/// Wall or floor formatting in 3D charts.
#[derive(Debug, Clone)]
pub struct WallFloor {
    /// Thickness (0-4096 points)
    pub thickness: Option<u32>,
}

impl WallFloor {
    /// Create a new wall/floor with default settings.
    #[inline]
    pub fn new() -> Self {
        Self { thickness: None }
    }

    /// Set thickness.
    #[inline]
    pub fn with_thickness(mut self, thickness: u32) -> Self {
        self.thickness = Some(thickness);
        self
    }
}

impl Default for WallFloor {
    #[inline]
    fn default() -> Self {
        Self::new()
    }
}

/// The main chart structure.
#[derive(Debug, Clone)]
pub struct Chart {
    /// Chart title
    pub title: Option<TitleText>,
    /// Manual layout for the chart title
    pub title_layout: Option<Layout>,
    /// Whether the chart title overlays the plot area
    pub title_overlay: bool,
    /// Whether auto-generated title has been deleted
    pub auto_title_deleted: bool,
    /// Optional pivot-chart formatting collection; `Some` preserves an empty wrapper
    pub pivot_formats: Option<Vec<PivotFormat>>,
    /// Plot area with series and axes
    pub plot_area: PlotArea,
    /// Chart legend
    pub legend: Option<Legend>,
    /// 3D view settings
    pub view_3d: Option<View3D>,
    /// Floor formatting (3D charts)
    pub floor: Option<WallFloor>,
    /// Back wall formatting (3D charts)
    pub back_wall: Option<WallFloor>,
    /// Side wall formatting (3D charts)
    pub side_wall: Option<WallFloor>,
    /// How to display blank values
    pub display_blanks_as: DisplayBlanks,
    /// Plot only visible cells
    pub plot_visible_only: bool,
    /// Show data in hidden rows and columns
    pub show_data_labels_over_max: bool,
    /// Chart style index
    pub style: Option<u32>,
    /// Chart content language
    pub language: Option<String>,
    /// Optional pivot-table source metadata
    pub pivot_source: Option<PivotSource>,
    /// Use 1904 date system
    pub date_1904: bool,
    /// Rounding corners
    pub rounded_corners: bool,
    /// Optional chart printing configuration
    pub print_settings: Option<ChartPrintSettings>,
}

impl Chart {
    /// Create a new chart with default settings.
    #[inline]
    pub fn new() -> Self {
        Self {
            title: None,
            title_layout: None,
            title_overlay: false,
            auto_title_deleted: false,
            pivot_formats: None,
            plot_area: PlotArea::new(),
            legend: None,
            view_3d: None,
            floor: None,
            back_wall: None,
            side_wall: None,
            display_blanks_as: DisplayBlanks::Gap,
            plot_visible_only: true,
            show_data_labels_over_max: false,
            style: None,
            language: None,
            pivot_source: None,
            date_1904: false,
            rounded_corners: false,
            print_settings: None,
        }
    }

    /// Set the chart title.
    #[inline]
    pub fn with_title(mut self, title: impl Into<String>) -> Self {
        self.title = Some(TitleText::from_string(title));
        self
    }

    /// Set the plot area.
    #[inline]
    pub fn with_plot_area(mut self, plot_area: PlotArea) -> Self {
        self.plot_area = plot_area;
        self
    }

    /// Set the legend.
    #[inline]
    pub fn with_legend(mut self, legend: Legend) -> Self {
        self.legend = Some(legend);
        self
    }

    /// Enable 3D view.
    #[inline]
    pub fn with_3d_view(mut self, view: View3D) -> Self {
        self.view_3d = Some(view);
        self
    }

    /// Check if this is a 3D chart.
    #[inline]
    pub fn is_3d(&self) -> bool {
        self.view_3d.is_some()
            || self.plot_area.type_groups.iter().any(|tg| {
                matches!(
                    tg,
                    crate::charts::plot_area::TypeGroup::Area3D(_)
                        | crate::charts::plot_area::TypeGroup::Bar3D(_)
                        | crate::charts::plot_area::TypeGroup::Line3D(_)
                        | crate::charts::plot_area::TypeGroup::Pie3D(_)
                        | crate::charts::plot_area::TypeGroup::Surface3D(_)
                )
            })
    }
}

impl Default for Chart {
    #[inline]
    fn default() -> Self {
        Self::new()
    }
}
