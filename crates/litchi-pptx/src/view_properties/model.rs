//! Package-independent `PresentationML` view-property values.

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ViewKind {
    Slide,
    SlideMaster,
    Notes,
    Handout,
    NotesMaster,
    Outline,
    SlideSorter,
    SlideThumbnail,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SplitterState {
    Minimized,
    Restored,
    Maximized,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum GuideOrientation {
    Horizontal,
    Vertical,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Ratio {
    pub numerator: i64,
    pub denominator: i64,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Point {
    pub x: i64,
    pub y: i64,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CommonView {
    pub variable_scale: Option<bool>,
    pub scale_x: Ratio,
    pub scale_y: Ratio,
    pub origin: Point,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Guide {
    pub orientation: Option<GuideOrientation>,
    pub position: Option<i32>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CommonSlideView {
    pub snap_to_grid: Option<bool>,
    pub snap_to_objects: Option<bool>,
    pub show_guides: Option<bool>,
    pub view: CommonView,
    pub guides: Vec<Guide>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct RestoredPane {
    pub size: u32,
    pub auto_adjust: Option<bool>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct NormalView {
    pub show_outline_icons: Option<bool>,
    pub snap_vertical_splitter: Option<bool>,
    pub vertical_bar_state: Option<SplitterState>,
    pub horizontal_bar_state: Option<SplitterState>,
    pub prefer_single_view: Option<bool>,
    pub restored_left: RestoredPane,
    pub restored_top: RestoredPane,
    pub extension_xml: Option<Vec<u8>>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OutlineSlide {
    pub relationship_id: String,
    pub collapse: Option<bool>,
    pub target: Option<String>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OutlineView {
    pub view: CommonView,
    pub slides: Vec<OutlineSlide>,
    pub extension_xml: Option<Vec<u8>>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SimpleView {
    pub view: CommonView,
    pub extension_xml: Option<Vec<u8>>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SorterView {
    pub show_formatting: Option<bool>,
    pub view: CommonView,
    pub extension_xml: Option<Vec<u8>>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SlideLikeView {
    pub common: CommonSlideView,
    pub extension_xml: Option<Vec<u8>>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct GridSpacing {
    pub cx: u32,
    pub cy: u32,
}

#[derive(Clone, Debug, PartialEq, Eq, Default)]
pub struct ViewProperties {
    pub last_view: Option<ViewKind>,
    pub show_comments: Option<bool>,
    pub normal: Option<NormalView>,
    pub slide: Option<SlideLikeView>,
    pub outline: Option<OutlineView>,
    pub notes_text: Option<SimpleView>,
    pub sorter: Option<SorterView>,
    pub notes: Option<SlideLikeView>,
    pub grid_spacing: Option<GridSpacing>,
    pub extension_xml: Option<Vec<u8>>,
}
