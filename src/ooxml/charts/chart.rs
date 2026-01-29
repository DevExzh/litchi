//! Main chart structure.
//!
//! This module contains the top-level chart structure that combines
//! all chart elements (plot area, legend, title, etc.).

use crate::ooxml::charts::legend::Legend;
use crate::ooxml::charts::models::TitleText;
use crate::ooxml::charts::plot_area::PlotArea;
use crate::ooxml::charts::types::DisplayBlanks;

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
    /// Whether auto-generated title has been deleted
    pub auto_title_deleted: bool,
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
    /// Use 1904 date system
    pub date_1904: bool,
    /// Rounding corners
    pub rounded_corners: bool,
}

impl Chart {
    /// Create a new chart with default settings.
    #[inline]
    pub fn new() -> Self {
        Self {
            title: None,
            auto_title_deleted: false,
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
            date_1904: false,
            rounded_corners: false,
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
                    crate::ooxml::charts::plot_area::TypeGroup::Area3D(_)
                        | crate::ooxml::charts::plot_area::TypeGroup::Bar3D(_)
                        | crate::ooxml::charts::plot_area::TypeGroup::Line3D(_)
                        | crate::ooxml::charts::plot_area::TypeGroup::Pie3D(_)
                        | crate::ooxml::charts::plot_area::TypeGroup::Surface3D(_)
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
