//! Main chart structure.
//!
//! This module contains the top-level chart structure that combines
//! all chart elements (plot area, legend, title, etc.).

use crate::chart::data::{Layout, TitleText};
use crate::chart::legend::Legend;
use crate::chart::plot_area::PlotArea;
use crate::chart::series::{DataLabel, Marker};
use crate::chart::types::DisplayBlanks;
use crate::{Error, Result};
use litchi_ooxml_common::xml::is_drawingml_chart_name;
use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

fn validate_chart_xml_fragment(xml: &[u8], expected_root: &[u8], description: &str) -> Result<()> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut saw_root = false;
    let mut closed_root = false;
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                if matches!(namespace, ResolveResult::Unknown(_)) {
                    return Err(Error::Invalid(format!(
                        "{description} contains an undeclared element prefix"
                    )));
                }
                let has_expected_root =
                    is_drawingml_chart_name(&namespace, element.name(), expected_root);
                drop(namespace);
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
                    if matches!(
                        reader.resolver().resolve_attribute(attribute.key).0,
                        ResolveResult::Unknown(_)
                    ) {
                        return Err(Error::Invalid(format!(
                            "{description} contains an undeclared attribute prefix"
                        )));
                    }
                }
                if depth == 0 {
                    if saw_root || !has_expected_root {
                        return Err(Error::Invalid(format!(
                            "{description} must have one DrawingML chart root"
                        )));
                    }
                    saw_root = true;
                }
                if matches!(event, Event::Start(_)) {
                    depth = depth.checked_add(1).ok_or_else(|| {
                        Error::Invalid(format!("{description} XML nesting is too deep"))
                    })?;
                } else if depth == 0 {
                    closed_root = true;
                }
            },
            Event::End(ref element) => {
                if matches!(namespace, ResolveResult::Unknown(_)) {
                    return Err(Error::Invalid(format!(
                        "{description} contains an undeclared closing-element prefix"
                    )));
                }
                let has_expected_root =
                    is_drawingml_chart_name(&namespace, element.name(), expected_root);
                drop(namespace);
                if depth == 0 {
                    return Err(Error::Invalid(format!(
                        "{description} has an unmatched closing element"
                    )));
                }
                depth -= 1;
                if depth == 0 {
                    if !has_expected_root {
                        return Err(Error::Invalid(format!(
                            "{description} has an invalid root closing element"
                        )));
                    }
                    closed_root = true;
                }
            },
            Event::Text(ref text)
                if depth == 0
                    && !text
                        .decode()
                        .map_err(|error| Error::Xml(error.to_string()))?
                        .trim()
                        .is_empty() =>
            {
                return Err(Error::Invalid(format!(
                    "{description} contains text outside its root"
                )));
            },
            Event::CData(_) | Event::GeneralRef(_) if depth == 0 => {
                return Err(Error::Invalid(format!(
                    "{description} contains data outside its root"
                )));
            },
            Event::Decl(_) | Event::DocType(_) => {
                return Err(Error::Invalid(format!(
                    "{description} cannot contain an XML declaration or document type"
                )));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
    if !saw_root || !closed_root || depth != 0 {
        return Err(Error::Invalid(format!(
            "{description} has no complete root"
        )));
    }
    Ok(())
}

macro_rules! chart_xml_fragment {
    ($(#[$meta:meta])* $name:ident, $root:literal, $description:literal) => {
        $(#[$meta])*
        #[derive(Debug, Clone, PartialEq, Eq)]
        pub struct $name {
            xml: Vec<u8>,
        }

        impl $name {
            /// Validate and retain one complete XML fragment.
            /// # Errors
            ///
            /// Returns an error when input violates DrawingML constraints, exceeds a configured
            /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
            pub fn from_xml(xml: Vec<u8>) -> Result<Self> {
                validate_chart_xml_fragment(&xml, $root, $description)?;
                Ok(Self { xml })
            }

            /// Return the complete XML fragment.
            pub fn as_xml(&self) -> &[u8] {
                &self.xml
            }
        }
    };
}

chart_xml_fragment!(
    /// Complete chart-space shape properties, including arbitrary DrawingML children.
    ShapeProperties,
    b"spPr",
    "chart shape properties"
);

chart_xml_fragment!(
    /// Complete chart-space text properties, including arbitrary DrawingML children.
    TextProperties,
    b"txPr",
    "chart text properties"
);

chart_xml_fragment!(
    /// Complete chart-space extension list, including extension namespace content.
    ExtensionList,
    b"extLst",
    "chart extension list"
);

/// Printed chart header and footer strings and selection flags.
#[derive(Debug, Clone)]
pub struct HeaderFooter {
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

impl HeaderFooter {
    /// Create an empty header/footer using schema defaults.
    #[inline]
    #[must_use]
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

impl Default for HeaderFooter {
    #[inline]
    fn default() -> Self {
        Self::new()
    }
}

/// Page margins for printing a chart, in inches.
#[derive(Debug, Clone, Copy)]
pub struct PageMargins {
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

impl PageMargins {
    /// Create a complete page-margin set.
    #[inline]
    #[must_use]
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
pub enum PageOrientation {
    /// Use the printer default
    #[default]
    Default,
    /// Portrait orientation
    Portrait,
    /// Landscape orientation
    Landscape,
}

impl PageOrientation {
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
pub struct PageSetup {
    /// Printer paper-size code
    pub paper_size: u32,
    /// First printed page number
    pub first_page_number: u32,
    /// Page orientation
    pub orientation: PageOrientation,
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

impl PageSetup {
    /// Create page setup using schema defaults.
    #[inline]
    #[must_use]
    pub fn new() -> Self {
        Self {
            paper_size: 1,
            first_page_number: 1,
            orientation: PageOrientation::Default,
            black_and_white: false,
            draft: false,
            use_first_page_number: false,
            horizontal_dpi: 600,
            vertical_dpi: 600,
            copies: 1,
        }
    }
}

impl Default for PageSetup {
    #[inline]
    fn default() -> Self {
        Self::new()
    }
}

/// Optional chart printing configuration.
#[derive(Debug, Clone, Default)]
pub struct PrintSettings {
    /// Header and footer settings
    pub header_footer: Option<HeaderFooter>,
    /// Page margins
    pub page_margins: Option<PageMargins>,
    /// Page setup
    pub page_setup: Option<PageSetup>,
}

impl PrintSettings {
    /// Create an explicitly empty print-settings container.
    #[inline]
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }
}

/// Formatting override for one pivot-chart data point.
#[derive(Debug, Clone)]
pub struct PivotFormat {
    /// Zero-based data-point index
    pub index: u32,
    /// Shape properties for the pivot-format point
    pub shape_properties: Option<ShapeProperties>,
    /// Text properties for the pivot-format point
    pub text_properties: Option<TextProperties>,
    /// Optional marker override
    pub marker: Option<Marker>,
    /// Optional data-label override
    pub data_label: Option<DataLabel>,
    /// Pivot-format extension list
    pub extension_list: Option<ExtensionList>,
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

/// Interaction restrictions applied when the containing sheet is protected.
///
/// Each switch is optional because an omitted switch is distinct from an
/// explicitly enabled or disabled switch in the chart XML.
#[derive(Debug, Clone, Default)]
pub struct Protection {
    /// Prevent changes to the chart object
    pub chart_object: Option<bool>,
    /// Prevent changes to chart data
    pub data: Option<bool>,
    /// Prevent chart formatting changes
    pub formatting: Option<bool>,
    /// Prevent chart-element selection
    pub selection: Option<bool>,
    /// Prevent chart user-interface interaction
    pub user_interface: Option<bool>,
}

/// A `DrawingML` theme color that can be selected by a color mapping.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ColorSchemeIndex {
    /// Dark theme color 1
    Dark1,
    /// Light theme color 1
    Light1,
    /// Dark theme color 2
    Dark2,
    /// Light theme color 2
    Light2,
    /// Accent theme color 1
    Accent1,
    /// Accent theme color 2
    Accent2,
    /// Accent theme color 3
    Accent3,
    /// Accent theme color 4
    Accent4,
    /// Accent theme color 5
    Accent5,
    /// Accent theme color 6
    Accent6,
    /// Hyperlink theme color
    Hyperlink,
    /// Followed-hyperlink theme color
    FollowedHyperlink,
}

impl ColorSchemeIndex {
    pub(crate) const fn as_str(self) -> &'static str {
        match self {
            Self::Dark1 => "dk1",
            Self::Light1 => "lt1",
            Self::Dark2 => "dk2",
            Self::Light2 => "lt2",
            Self::Accent1 => "accent1",
            Self::Accent2 => "accent2",
            Self::Accent3 => "accent3",
            Self::Accent4 => "accent4",
            Self::Accent5 => "accent5",
            Self::Accent6 => "accent6",
            Self::Hyperlink => "hlink",
            Self::FollowedHyperlink => "folHlink",
        }
    }
}

/// Complete `DrawingML` theme color mapping used by a chart override.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ColorMapping {
    /// Background color 1 mapping
    pub background1: ColorSchemeIndex,
    /// Text color 1 mapping
    pub text1: ColorSchemeIndex,
    /// Background color 2 mapping
    pub background2: ColorSchemeIndex,
    /// Text color 2 mapping
    pub text2: ColorSchemeIndex,
    /// Accent color 1 mapping
    pub accent1: ColorSchemeIndex,
    /// Accent color 2 mapping
    pub accent2: ColorSchemeIndex,
    /// Accent color 3 mapping
    pub accent3: ColorSchemeIndex,
    /// Accent color 4 mapping
    pub accent4: ColorSchemeIndex,
    /// Accent color 5 mapping
    pub accent5: ColorSchemeIndex,
    /// Accent color 6 mapping
    pub accent6: ColorSchemeIndex,
    /// Hyperlink color mapping
    pub hyperlink: ColorSchemeIndex,
    /// Followed-hyperlink color mapping
    pub followed_hyperlink: ColorSchemeIndex,
}

impl Default for ColorMapping {
    fn default() -> Self {
        Self {
            background1: ColorSchemeIndex::Light1,
            text1: ColorSchemeIndex::Dark1,
            background2: ColorSchemeIndex::Light2,
            text2: ColorSchemeIndex::Dark2,
            accent1: ColorSchemeIndex::Accent1,
            accent2: ColorSchemeIndex::Accent2,
            accent3: ColorSchemeIndex::Accent3,
            accent4: ColorSchemeIndex::Accent4,
            accent5: ColorSchemeIndex::Accent5,
            accent6: ColorSchemeIndex::Accent6,
            hyperlink: ColorSchemeIndex::Hyperlink,
            followed_hyperlink: ColorSchemeIndex::FollowedHyperlink,
        }
    }
}

/// Chart color-map override selection.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ColorMapOverride {
    /// Inherit the color mapping from the chart's master context
    Master,
    /// Use the supplied complete color mapping
    Override(ColorMapping),
}

/// Relationship metadata for the data source embedded in or linked from a chart.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalData {
    /// Relationship identifier in the chart part, if already allocated
    pub relationship_id: Option<String>,
    /// Automatic-update setting; `None` preserves an omitted child element
    pub auto_update: Option<bool>,
}

/// Relationship metadata for a chart user-shapes drawing part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UserShapes {
    /// Relationship identifier in the chart part, if already allocated
    pub relationship_id: Option<String>,
}

impl UserShapes {
    /// Create user-shapes metadata for an existing chart relationship.
    pub fn new(relationship_id: impl Into<String>) -> Self {
        Self {
            relationship_id: Some(relationship_id.into()),
        }
    }

    /// Create metadata whose relationship will be allocated by a package writer.
    #[must_use]
    pub fn pending() -> Self {
        Self {
            relationship_id: None,
        }
    }
}

impl ExternalData {
    /// Create external-data metadata for an existing chart relationship.
    pub fn new(relationship_id: impl Into<String>) -> Self {
        Self {
            relationship_id: Some(relationship_id.into()),
            auto_update: None,
        }
    }

    /// Create metadata whose relationship will be allocated by a package writer.
    #[must_use]
    pub fn pending() -> Self {
        Self {
            relationship_id: None,
            auto_update: None,
        }
    }
}

impl PivotFormat {
    /// Create a pivot-format entry for one data point.
    #[inline]
    #[must_use]
    pub fn new(index: u32) -> Self {
        Self {
            index,
            shape_properties: None,
            text_properties: None,
            marker: None,
            data_label: None,
            extension_list: None,
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
    #[must_use]
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
    #[must_use]
    pub fn with_rotation(mut self, rot_x: u32, rot_y: u32) -> Self {
        self.rot_x = Some(rot_x);
        self.rot_y = Some(rot_y);
        self
    }

    /// Set perspective.
    #[inline]
    #[must_use]
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

/// How a picture fill is applied to a chart surface.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PictureFormat {
    /// Stretch one picture across the surface
    Stretch,
    /// Tile pictures at their natural scale
    Stack,
    /// Tile pictures using an explicit stack unit
    StackScale,
}

impl PictureFormat {
    pub(crate) fn xml_value(self) -> &'static str {
        match self {
            Self::Stretch => "stretch",
            Self::Stack => "stack",
            Self::StackScale => "stackScale",
        }
    }
}

/// Picture-fill placement options used by chart surfaces and data points.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct PictureOptions {
    /// Apply the picture to the front face
    pub apply_to_front: Option<bool>,
    /// Apply the picture to side faces
    pub apply_to_sides: Option<bool>,
    /// Apply the picture to the end face
    pub apply_to_end: Option<bool>,
    /// Stretch or stacking behavior
    pub picture_format: Option<PictureFormat>,
    /// Number of picture units represented by each stacked picture
    pub picture_stack_unit: Option<f64>,
}

/// Wall or floor formatting in 3D charts.
#[derive(Debug, Clone)]
pub struct WallFloor {
    /// Thickness (0-4096 points)
    pub thickness: Option<u32>,
    /// `DrawingML` shape properties for the surface
    pub shape_properties: Option<ShapeProperties>,
    /// Picture-fill placement options
    pub picture_options: Option<PictureOptions>,
    /// Surface extension list
    pub extension_list: Option<ExtensionList>,
}

impl WallFloor {
    /// Create a new wall/floor with default settings.
    #[inline]
    #[must_use]
    pub fn new() -> Self {
        Self {
            thickness: None,
            shape_properties: None,
            picture_options: None,
            extension_list: None,
        }
    }

    /// Set thickness.
    #[inline]
    #[must_use]
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
    /// `DrawingML` shape properties for the chart title
    pub title_shape_properties: Option<ShapeProperties>,
    /// `DrawingML` text properties for the chart title
    pub title_text_properties: Option<TextProperties>,
    /// Chart-title extension list
    pub title_extension_list: Option<ExtensionList>,
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
    /// Optional extension list inside the chart element
    pub chart_extension_list: Option<ExtensionList>,
    /// Chart style index
    pub style: Option<u32>,
    /// Optional `DrawingML` theme color-map override
    pub color_map_override: Option<ColorMapOverride>,
    /// Chart content language
    pub language: Option<String>,
    /// Optional pivot-table source metadata
    pub pivot_source: Option<PivotSource>,
    /// Optional chart interaction protection; `Some` preserves an empty wrapper
    pub protection: Option<Protection>,
    /// Use 1904 date system
    pub date_1904: bool,
    /// Rounding corners
    pub rounded_corners: bool,
    /// Optional external or embedded chart data relationship metadata
    pub external_data: Option<ExternalData>,
    /// Optional chart user-shapes drawing relationship metadata
    pub user_shapes: Option<UserShapes>,
    /// Optional chart-space `DrawingML` shape properties
    pub shape_properties: Option<ShapeProperties>,
    /// Optional chart-space `DrawingML` text properties
    pub text_properties: Option<TextProperties>,
    /// Optional chart printing configuration
    pub print_settings: Option<PrintSettings>,
    /// Optional chart-space extension list
    pub extension_list: Option<ExtensionList>,
}

impl Chart {
    /// Create a new chart with default settings.
    #[inline]
    #[must_use]
    pub fn new() -> Self {
        Self {
            title: None,
            title_layout: None,
            title_overlay: false,
            title_shape_properties: None,
            title_text_properties: None,
            title_extension_list: None,
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
            chart_extension_list: None,
            style: None,
            color_map_override: None,
            language: None,
            pivot_source: None,
            protection: None,
            date_1904: false,
            rounded_corners: false,
            external_data: None,
            user_shapes: None,
            shape_properties: None,
            text_properties: None,
            print_settings: None,
            extension_list: None,
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
    #[must_use]
    pub fn with_plot_area(mut self, plot_area: PlotArea) -> Self {
        self.plot_area = plot_area;
        self
    }

    /// Set the legend.
    #[inline]
    #[must_use]
    pub fn with_legend(mut self, legend: Legend) -> Self {
        self.legend = Some(legend);
        self
    }

    /// Enable 3D view.
    #[inline]
    #[must_use]
    pub fn with_3d_view(mut self, view: View3D) -> Self {
        self.view_3d = Some(view);
        self
    }

    /// Check if this is a 3D chart.
    #[inline]
    #[must_use]
    pub fn is_3d(&self) -> bool {
        self.view_3d.is_some()
            || self.plot_area.type_groups.iter().any(|tg| {
                matches!(
                    tg,
                    crate::chart::plot_area::TypeGroup::Area3D(_)
                        | crate::chart::plot_area::TypeGroup::Bar3D(_)
                        | crate::chart::plot_area::TypeGroup::Line3D(_)
                        | crate::chart::plot_area::TypeGroup::Pie3D(_)
                        | crate::chart::plot_area::TypeGroup::Surface3D(_)
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
